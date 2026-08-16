/**
 * Power Query section documents.
 *
 * The extension reads and writes ONE unit on disk: a section document holding every query in the
 * workbook.
 *
 *     section Section1;
 *
 *     shared fGetNamedRange = let ... ;
 *     shared RawInput       = let ... ;
 *     shared FinalTable     = let ... ;
 *
 * Excel's object model works in a DIFFERENT unit. `WorkbookQuery.Formula` is one query's
 * expression, with no `shared` and no trailing semicolon. Measured: 160 characters written through
 * COM came back from a saved file as 199 — the delta being exactly the section header, the
 * `shared <name> = ` prefix and the closing `;`.
 *
 * So writing live cannot hand the file's text to Excel. The document has to be split by name, and
 * each expression written to its matching query. This module is that split, and its inverse.
 *
 * IT IS DELIBERATELY NOT AN M PARSER. It finds top-level `shared` bindings and slices the text
 * between them. Anything inside an expression - nested `let`, comments, strings containing the word
 * shared - is carried across verbatim, because the safest transformation of somebody's code is the
 * one that does not understand it.
 */

export interface MQuery {
	/** Query name as it appears after `shared`, unquoted. */
	name: string;
	/** The expression exactly as written, with no trailing semicolon or whitespace. */
	expression: string;
	/**
	 * Whitespace that sat between the expression and its `;`.
	 *
	 * Kept so a rebuild is lossless: real documents write `AddedDate
;` as often as
	 * `AddedDate;`, and a round trip that "tidies" somebody's formatting is drift by another name.
	 * It is deliberately NOT part of `expression`, because `expression` is what goes to Excel.
	 */
	terminator: string;
	/** True when the source wrote the name in #"quoted" form. */
	quoted: boolean;
	/**
	 * Everything between this binding's `;` and the next binding - or the end of the document.
	 *
	 * Usually a blank line. Sometimes a comment ABOUT the next query, which belongs to the document
	 * and not to either expression. Dropping it on rebuild would be silent drift in somebody's file,
	 * so it is carried verbatim. Absent when a section was built by hand rather than parsed.
	 */
	separator?: string;
}

export interface MSection {
	/** Text before the first `shared` binding - the `section` line and anything above it. */
	header: string;
	queries: MQuery[];
}

/**
 * `shared Foo =` / `shared #"Foo Bar" =`, matched only where the scanner says we are in code.
 *
 * Sticky, because it is tried at one position rather than searched for.
 */
const BINDING_AT = /shared[ \t]+(?:#"((?:[^"]|"")*)"|([A-Za-z_][\w.]*))[ \t]*=/y;

interface FoundBinding {
	/** Start of the line the binding sits on, so leading indentation belongs to it. */
	start: number;
	/** Just past the `=`. */
	exprStart: number;
	/** The terminating `;`, or -1 if the document ends without one. */
	semi: number;
	name: string;
	quoted: boolean;
}

/**
 * Find top-level `shared` bindings, ignoring anything inside a string or a comment.
 *
 * THIS IS NOT AN M PARSER, AND IS NOT BECOMING ONE. It tracks four states - code, string, line
 * comment, block comment - and does nothing else. That is the minimum needed to stop mistaking an
 * expression's CONTENTS for the document's STRUCTURE, which a regex over lines cannot help doing:
 *
 *     shared Real = let
 *         x = 1,
 *         /*
 *     shared Fake = 666        <- starts a line, is not a binding
 *         *\/
 *         y = 2
 *     in y;
 *
 * The terminator is found the same way: the FIRST `;` seen in code after the expression begins.
 * Assuming instead that it is the last non-whitespace character before the next `shared` is wrong
 * the moment a comment sits between two bindings, and quietly folds that comment into the previous
 * expression.
 *
 * `#"quoted names"` are scanned as strings - they use the same `""` escape - so a `;` or the word
 * `shared` inside one cannot be mistaken for structure either.
 */
function scanBindings(document: string): FoundBinding[] {
	const found: FoundBinding[] = [];
	let state: 'code' | 'string' | 'line' | 'block' = 'code';
	let lineStart = 0;
	let onlyBlankSoFar = true;
	let i = 0;

	const closeSemi = (at: number): void => {
		for (let k = found.length - 1; k >= 0; k--) {
			if (found[k].semi === -1) { found[k].semi = at; return; }
		}
	};

	while (i < document.length) {
		const c = document[i];
		const next = document[i + 1];

		if (state === 'code') {
			if (c === '\n') { lineStart = i + 1; onlyBlankSoFar = true; i++; continue; }
			if (c === '/' && next === '/') { state = 'line'; i += 2; continue; }
			if (c === '/' && next === '*') { state = 'block'; i += 2; continue; }
			if (c === '#' && next === '"') { state = 'string'; i += 2; continue; }
			if (c === '"') { state = 'string'; i++; continue; }
			if (c === ';') { closeSemi(i); onlyBlankSoFar = false; i++; continue; }

			// A binding only counts at the start of a line, which is how section documents are
			// written and keeps `shared` inside an expression from being treated as structure.
			if (onlyBlankSoFar && c !== ' ' && c !== '\t' && c !== '\r') {
				BINDING_AT.lastIndex = i;
				const m = BINDING_AT.exec(document);
				if (m !== null && m.index === i) {
					found.push({
						start: lineStart,
						exprStart: i + m[0].length,
						semi: -1,
						// `""` is how a literal quote is escaped inside #"..."
						name: m[1] !== undefined ? m[1].replace(/""/g, '"') : m[2],
						quoted: m[1] !== undefined
					});
					onlyBlankSoFar = false;
					i += m[0].length;
					continue;
				}
			}

			if (c !== ' ' && c !== '\t' && c !== '\r') { onlyBlankSoFar = false; }
			i++;
			continue;
		}

		if (state === 'string') {
			if (c === '"') {
				if (next === '"') { i += 2; continue; }   // escaped quote, still inside
				state = 'code';
			}
			i++;
			continue;
		}

		if (state === 'line') {
			if (c === '\n') { state = 'code'; lineStart = i + 1; onlyBlankSoFar = true; }
			i++;
			continue;
		}

		// block comment - M's block comments do not nest
		if (c === '*' && next === '/') { state = 'code'; i += 2; continue; }
		if (c === '\n') { lineStart = i + 1; }
		i++;
	}

	return found;
}

/**
 * Split a section document into its queries.
 *
 * Returns `queries: []` for anything that has no `shared` bindings, which is the honest answer for
 * an empty or unrecognized document - the caller decides whether that is an error.
 */
export function parseSection(document: string): MSection {
	const found = scanBindings(document);
	if (found.length === 0) {
		return { header: document, queries: [] };
	}

	const queries: MQuery[] = [];
	for (let i = 0; i < found.length; i++) {
		const b = found[i];
		const nextStart = i + 1 < found.length ? found[i + 1].start : document.length;
		const exprEnd = b.semi >= 0 ? b.semi : nextStart;

		const raw = document.slice(b.exprStart, exprEnd);
		const lead = /^[ \t]*\r?\n?/.exec(raw)![0];
		const body = raw.slice(lead.length);
		const expression = body.replace(/\s+$/, '');

		queries.push({
			name: b.name,
			expression,
			terminator: body.slice(expression.length),
			quoted: b.quoted,
			separator: b.semi >= 0 ? document.slice(b.semi + 1, nextStart) : ''
		});
	}

	return { header: document.slice(0, found[0].start), queries };
}

/** Quote a name back the way M requires it. */
export function formatName(name: string, quoted: boolean): string {
	const needsQuoting = quoted || !/^[A-Za-z_][\w.]*$/.test(name);
	return needsQuoting ? `#"${name.replace(/"/g, '""')}"` : name;
}

/**
 * Rebuild a section document from its parts.
 *
 * `eol` defaults to CRLF because that is what Excel stores and what the on-disk writer produces;
 * passing the document's own line ending keeps a round trip honest.
 */
export function buildSection(section: MSection, eol: string = '\r\n'): string {
	// A parsed section carries the exact text that sat between its bindings, including any comments,
	// and replaying it is what makes the round trip byte-for-byte. A section assembled in code has
	// no separators, so fall back to the conventional blank line between queries.
	const parsed = section.queries.length > 0 && section.queries.every(q => q.separator !== undefined);

	if (parsed) {
		const header = section.header;
		const body = section.queries
			.map(q => `shared ${formatName(q.name, q.quoted)} = ${q.expression}${q.terminator};${q.separator}`)
			.join('');
		return header + body;
	}

	const header = section.header.replace(/\s+$/, '');
	const body = section.queries
		.map(q => `shared ${formatName(q.name, q.quoted)} = ${q.expression}${q.terminator};`)
		.join(eol + eol);
	return header + eol + eol + body + eol;
}

/** Whichever line ending the document actually uses. Ties go to CRLF, Excel's own. */
export function detectEol(document: string): string {
	const crlf = (document.match(/\r\n/g) ?? []).length;
	const lf = (document.match(/(?<!\r)\n/g) ?? []).length;
	return lf > crlf ? '\n' : '\r\n';
}

export interface QueryDiff {
	/** In the document and in the workbook - write the expression. */
	update: MQuery[];
	/** In the document but not the workbook - the user added a query. */
	add: MQuery[];
	/** In the workbook but not the document. NEVER acted on automatically. */
	missingFromDocument: string[];
}

/**
 * Work out what a live sync would have to do.
 *
 * `missingFromDocument` is reported and never actioned. A query in the workbook with no matching
 * binding might mean the user deleted it in the editor - or that they are syncing a document
 * extracted before that query existed. Deleting somebody's query on a guess is not a trade this
 * project makes; the caller surfaces it and lets a human decide.
 */
export function diffQueries(document: MSection, workbookQueryNames: string[]): QueryDiff {
	const inWorkbook = new Set(workbookQueryNames);
	const inDocument = new Set(document.queries.map(q => q.name));

	return {
		update: document.queries.filter(q => inWorkbook.has(q.name)),
		add: document.queries.filter(q => !inWorkbook.has(q.name)),
		missingFromDocument: workbookQueryNames.filter(n => !inDocument.has(n))
	};
}
