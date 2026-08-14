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
}

export interface MSection {
	/** Text before the first `shared` binding - the `section` line and anything above it. */
	header: string;
	queries: MQuery[];
}

/** `shared Foo =` or `shared #"Foo Bar" =` at the start of a line. */
const SHARED_BINDING = /^[ \t]*shared[ \t]+(?:#"((?:[^"]|"")*)"|([A-Za-z_][\w.]*))[ \t]*=/gm;

/**
 * Split a section document into its queries.
 *
 * Returns `queries: []` for anything that has no `shared` bindings, which is the honest answer for
 * an empty or unrecognised document - the caller decides whether that is an error.
 */
export function parseSection(document: string): MSection {
	const matches: { index: number; end: number; name: string; quoted: boolean }[] = [];

	SHARED_BINDING.lastIndex = 0;
	let m: RegExpExecArray | null;
	while ((m = SHARED_BINDING.exec(document)) !== null) {
		const quoted = m[1] !== undefined;
		matches.push({
			index: m.index,
			end: m.index + m[0].length,
			// `""` is how a literal quote is escaped inside #"..."
			name: quoted ? m[1].replace(/""/g, '"') : m[2],
			quoted
		});
	}

	if (matches.length === 0) {
		return { header: document, queries: [] };
	}

	const queries: MQuery[] = [];
	for (let i = 0; i < matches.length; i++) {
		const from = matches[i].end;
		const to = i + 1 < matches.length ? matches[i + 1].index : document.length;
		const binding = splitBinding(document.slice(from, to));
		queries.push({
			name: matches[i].name,
			expression: binding.expression,
			terminator: binding.terminator,
			quoted: matches[i].quoted
		});
	}

	return { header: document.slice(0, matches[0].index), queries };
}

/**
 * Split a binding's text into the expression and the whitespace before its `;`.
 *
 * Only ONE trailing semicolon is removed, and only if it is the last non-whitespace character. An
 * expression legitimately ending in `;` inside a string or comment is left alone, because the
 * semicolon we care about is the binding terminator and always sits outermost.
 */
function splitBinding(raw: string): { expression: string; terminator: string } {
	const text = raw.replace(/^[ \t]*\r?\n?/, '').replace(/\s+$/, '');
	if (!text.endsWith(';')) {
		return { expression: text, terminator: '' };
	}
	const withoutSemi = text.slice(0, -1);
	const expression = withoutSemi.replace(/\s+$/, '');
	return { expression, terminator: withoutSemi.slice(expression.length) };
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
