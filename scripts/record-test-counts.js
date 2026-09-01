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
		// SUM the summaries, do not take the last one. There is more than one test host now - a
		// default one and a second with a workspace folder open - so mocha prints a summary per
		// config, and reading only the last would report the size of whichever ran last.
		//
		// The patterns are anchored to a whole line so fixture logs containing the word "passing"
		// cannot be mistaken for a summary.
		const re = new RegExp(`^\\s*(\\d+)\\s+${word}\\b.*$`, 'gm');
		return [...output.matchAll(re)].reduce((total, m) => total + Number(m[1]), 0);
	};

	const passing = read('passing');
	const pending = read('pending');
	const failing = read('failing');
	const total = passing + pending + failing;

	if (!total) {
		console.error('\nrecord-test-counts: could not find a mocha summary in the output.');
		process.exit(code || 1);
	}

	// A FAILED RUN MUST NOT REWRITE THE COUNT.
	//
	// There is more than one test host. If the first dies - a compile error, a host that never
	// starts - the second can still finish and print a perfectly valid summary, and summing what
	// survived produces a number that looks real and is not. That happened: one host failed, the
	// other reported 5, and {"total": 5} was written over a file that said 136.
	//
	// The count only means anything when the whole suite ran, so a non-zero exit leaves the file
	// alone. --check bails here too: comparing against a partial count reports a mismatch and sends
	// somebody chasing a number when the real problem is the failure above it.
	if (code !== 0) {
		console.error('\nTests failed (exit ' + code + '); test-counts.json left unchanged. The partial run counted ' + total + '.');
		process.exit(code);
	}

	// ONLY `total` is written. passing/pending depend on whether the machine has Excel - this rig
	// reports 119 passing and 0 pending, every CI runner reports 114 and 5 - so persisting them makes
	// the file differ by machine, and any check comparing it fails for everyone but its author.
	// `total` is the same everywhere, which is the whole reason it is the number we publish.
	const next = JSON.stringify({ total }, null, 2) + '\n';
	const previous = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
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
