#!/usr/bin/env node
// Run the test suite and record what it actually contained.
//
// WHY NOT COUNT THE SOURCE. Two `test()` call sites sit inside `for` loops over three fixtures each,
// so a regex counting declarations reports 115 where the suite has 119. Counting text gets the wrong
// answer for any suite that generates tests, and this one does.
//
// WHY TOTAL, NOT PASSING. `passing` depends on the machine: the five live sync integration tests need
// Excel, so a developer rig with Excel reports 119 passing and 0 pending, while CI - on Windows as
// well as Linux, since no runner has Excel - reports 114 passing and 5 pending. **Total is the same
// everywhere**, which makes it the only count that is true regardless of where it was measured.
//
//   npm run test:counts     run the suite and rewrite test-counts.json
//
// docs/CONTRIBUTING.md reads these through .ewc3-docs.json, so `npm run docs:values` turns them into
// prose. CI verifies the recorded numbers still match a real run.

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, '..', 'test-counts.json');

const check = process.argv.includes('--check');

// The same command `npm test` runs. Kept here rather than shelling back through npm, which would
// recurse if npm test ever delegates to this script.
const child = spawn('npx', ['vscode-test'], {
	shell: true,
	stdio: ['inherit', 'pipe', 'inherit']
});

let output = '';
child.stdout.on('data', chunk => {
	const text = chunk.toString();
	output += text;
	process.stdout.write(text);
});

child.on('close', code => {
	const read = word => {
		// Mocha's summary. Take the LAST match: the fixture logs contain the word "passing" too.
		const matches = [...output.matchAll(new RegExp(`(\\d+)\\s+${word}\\b`, 'g'))];
		return matches.length ? Number(matches[matches.length - 1][1]) : 0;
	};

	const passing = read('passing');
	const pending = read('pending');
	const failing = read('failing');
	const total = passing + pending + failing;

	if (!total) {
		console.error('\nrecord-test-counts: could not find a mocha summary in the output.');
		process.exit(code || 1);
	}

	const counts = { total, passing, pending, failing };
	const next = JSON.stringify(counts, null, 2) + '\n';
	const previous = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';

	// Only `total` is portable. Comparing passing/pending across machines would fail on any rig whose
	// Excel situation differs from whoever last recorded them.
	const recordedTotal = previous ? (JSON.parse(previous).total || 0) : 0;

	if (check) {
		if (recordedTotal !== total) {
			console.error(`\ntest-counts.json says ${recordedTotal} tests; the suite has ${total}.`);
			console.error('Run: npm run test:counts && npm run docs:values');
			process.exit(1);
		}
		console.log(`\ntest-counts.json is current (${total} tests).`);
		process.exit(code);
	}

	if (previous !== next) {
		fs.writeFileSync(OUT, next);
		console.log(`\nRecorded ${total} tests (${passing} passing, ${pending} pending here).`);
	} else {
		console.log(`\ntest-counts.json unchanged (${total} tests).`);
	}
	process.exit(code);
});
