import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import { parseSection, buildSection, detectEol, diffQueries, formatName } from '../src/mSection';

/**
 * The section document is the unit the extension reads and writes on disk. Excel's object model
 * works one query at a time. Live sync stands or falls on translating between the two without
 * changing a byte of anyone's M, so these tests are mostly about losslessness.
 */
suite('M section documents', () => {

	const fixture = (name: string) =>
		fs.readFileSync(path.join(__dirname, '..', '..', 'test', 'fixtures', 'expected', name), 'utf8');

	suite('parsing', () => {
		test('finds every query in a multi-query document', () => {
			const sec = parseSection(fixture('complex_FinalTable.m'));
			assert.deepStrictEqual(sec.queries.map(q => q.name),
				['fGetNamedRange', 'RawInput', 'FinalTable']);
		});

		test('keeps the header, including extraction comments', () => {
			const sec = parseSection(fixture('simple_StudentResults.m'));
			assert.ok(sec.header.includes('section Section1;'), 'section line belongs to the header');
			assert.ok(!sec.header.includes('shared '), 'header stops at the first binding');
		});

		test('an expression carries no trailing semicolon - that is what Excel wants', () => {
			const sec = parseSection(fixture('simple_StudentResults.m'));
			const expr = sec.queries[0].expression;
			assert.ok(!expr.trimEnd().endsWith(';'), 'expression must not end in the binding terminator');
			assert.ok(expr.startsWith('let'), 'expression starts at the M, not the whitespace');
		});

		test('a document with no bindings is not an error', () => {
			const sec = parseSection('// just a comment\r\nsection Section1;\r\n');
			assert.strictEqual(sec.queries.length, 0);
			assert.ok(sec.header.length > 0);
		});

		test('reads #"quoted names" and unescapes them', () => {
			const sec = parseSection('section Section1;\r\n\r\nshared #"Sales By Region" = let x = 1 in x;\r\n');
			assert.strictEqual(sec.queries[0].name, 'Sales By Region');
			assert.strictEqual(sec.queries[0].quoted, true);
		});

		test('does not mistake the word shared inside an expression for a binding', () => {
			const doc = 'section Section1;\r\n\r\nshared Real = let\r\n    note = "shared Fake = 1",\r\n' +
				'    // shared AlsoFake = 2\r\n    out = note\r\nin\r\n    out;\r\n';
			const sec = parseSection(doc);
			assert.deepStrictEqual(sec.queries.map(q => q.name), ['Real'],
				'only the top-level binding counts');
		});
	});

	suite('lexical state - expression contents are not structure', () => {
		test('a shared binding inside a BLOCK COMMENT is not a query', () => {
			const doc = [
				'section Section1;',
				'',
				'shared Real = let',
				'    x = 1,',
				'    /*',
				'shared Fake = 666;',
				'    */',
				'    y = 2',
				'in',
				'    y;',
				''
			].join('\n');

			const sec = parseSection(doc);
			assert.deepStrictEqual(sec.queries.map(q => q.name), ['Real'],
				'a commented-out binding must not become a workbook query');
			assert.ok(sec.queries[0].expression.includes('shared Fake'),
				'and it must still be carried inside the expression, verbatim');
		});

		test('a shared binding inside a STRING is not a query', () => {
			const doc = [
				'section Section1;',
				'',
				'shared Real = let',
				'    note = "',
				'shared Fake = 666;',
				'",',
				'    y = 2',
				'in',
				'    y;',
				''
			].join('\n');

			assert.deepStrictEqual(parseSection(doc).queries.map(q => q.name), ['Real']);
		});

		test('a shared binding inside a LINE COMMENT is not a query', () => {
			const doc = 'section Section1;\n\nshared Real = 1;\n// shared Fake = 2;\n';
			assert.deepStrictEqual(parseSection(doc).queries.map(q => q.name), ['Real']);
		});

		test('a semicolon inside a string does not terminate the binding', () => {
			const doc = 'section Section1;\n\nshared Real = "a; b";\n';
			const sec = parseSection(doc);
			assert.strictEqual(sec.queries[0].expression, '"a; b"');
		});

		test('a quoted #"name" containing a semicolon is handled', () => {
			const doc = 'section Section1;\n\nshared #"odd; name" = 1;\n';
			const sec = parseSection(doc);
			assert.deepStrictEqual(sec.queries.map(q => q.name), ['odd; name']);
			assert.strictEqual(sec.queries[0].expression, '1');
		});
	});

	suite('comments between bindings', () => {
		const DOC = [
			'section Section1;',
			'',
			'shared A = 1;',
			'',
			'// here is why B is weird',
			'shared B = 2;',
			''
		].join('\n');

		test('the comment is NOT folded into the previous expression', () => {
			const sec = parseSection(DOC);
			assert.strictEqual(sec.queries[0].expression, '1',
				'A ended at its own semicolon, not at the next binding');
		});

		test('both queries are found', () => {
			assert.deepStrictEqual(parseSection(DOC).queries.map(q => q.name), ['A', 'B']);
		});

		test('the comment survives a round trip', () => {
			const rebuilt = buildSection(parseSection(DOC), detectEol(DOC));
			assert.strictEqual(rebuilt, DOC, 'a comment about a query belongs to the document');
		});
	});

	suite('round trips - the invariant live sync depends on', () => {
		for (const name of ['simple_StudentResults.m', 'complex_FinalTable.m', 'binary_FinalTable.m']) {
			test(`${name} rebuilds byte for byte`, () => {
				const doc = fixture(name);
				assert.strictEqual(buildSection(parseSection(doc), detectEol(doc)), doc);
			});
		}

		test('a semicolon on its own line stays on its own line', () => {
			// Real extracted documents write it both ways. Tidying somebody's formatting is drift.
			const doc = 'section Section1;\r\n\r\nshared A = let x = 1 in x\r\n;\r\n';
			assert.strictEqual(buildSection(parseSection(doc), '\r\n'), doc);
		});

		test('LF documents stay LF', () => {
			const doc = 'section Section1;\n\nshared A = let x = 1 in x;\n';
			assert.strictEqual(detectEol(doc), '\n');
			assert.strictEqual(buildSection(parseSection(doc), '\n'), doc);
		});

		test('a quoted name survives the trip', () => {
			const doc = 'section Section1;\r\n\r\nshared #"Sales By Region" = let x = 1 in x;\r\n';
			assert.strictEqual(buildSection(parseSection(doc), '\r\n'), doc);
		});
	});

	suite('names', () => {
		test('quotes a name that needs it', () => {
			assert.strictEqual(formatName('Sales By Region', false), '#"Sales By Region"');
			assert.strictEqual(formatName('Plain', false), 'Plain');
		});

		test('keeps quoting a name that was quoted, even if it need not be', () => {
			assert.strictEqual(formatName('Plain', true), '#"Plain"');
		});
	});

	suite('diffing against a workbook', () => {
		const doc = parseSection(
			'section Section1;\r\n\r\nshared A = let x = 1 in x;\r\n\r\nshared B = let x = 2 in x;\r\n');

		test('matches by name for updates', () => {
			const d = diffQueries(doc, ['A', 'B']);
			assert.deepStrictEqual(d.update.map(q => q.name), ['A', 'B']);
			assert.strictEqual(d.add.length, 0);
			assert.strictEqual(d.missingFromDocument.length, 0);
		});

		test('a binding with no query is an add', () => {
			const d = diffQueries(doc, ['A']);
			assert.deepStrictEqual(d.add.map(q => q.name), ['B']);
		});

		test('a query with no binding is REPORTED, never deleted', () => {
			const d = diffQueries(doc, ['A', 'B', 'Orphan']);
			assert.deepStrictEqual(d.missingFromDocument, ['Orphan']);
			// The point: nothing in the diff tells a caller to remove it.
			assert.ok(!('remove' in d), 'there is deliberately no remove list');
		});
	});
});
