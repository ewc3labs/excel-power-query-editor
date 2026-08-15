#!/usr/bin/env node
// Verify that every relative link in the docs points at something that exists.
//
// WHY. Two READMEs drifted until both source files were empty and nobody noticed for a year. A doc
// restructure just repointed six links that had been dangling. Neither failure was hard to find -
// they were only hard to NOTICE, because nothing was looking.
//
// This is the small version of the cross-repo route-and-usage mapping we use elsewhere: enumerate
// the references, resolve each one, report what does not land. One repo, no configuration.
//
//   npm run docs:links

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SKIP_DIRS = new Set(['node_modules', '.git', 'dist', 'out', '.vscode-test', 'archive']);

/** Markdown links and images, minus code spans. */
const LINK = /!?\[[^\]]*\]\(([^)]+)\)/g;
const FENCE = /```[\s\S]*?```|`[^`\n]*`/g;

function markdownFiles(dir, found = []) {
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		if (entry.isDirectory()) {
			if (!SKIP_DIRS.has(entry.name)) { markdownFiles(path.join(dir, entry.name), found); }
		} else if (entry.name.endsWith('.md')) {
			found.push(path.join(dir, entry.name));
		}
	}
	return found;
}

function isExternal(target) {
	return /^(https?:|mailto:|#)/i.test(target);
}

/**
 * Does this path exist with exactly the case that was written?
 *
 * fs.existsSync says yes on a case-insensitive filesystem regardless, so walk the path and confirm
 * each segment against the real directory listing.
 */
function existsWithExactCase(target) {
	let current = target;
	while (current !== ROOT && current !== path.dirname(current)) {
		const parent = path.dirname(current);
		if (!fs.readdirSync(parent).includes(path.basename(current))) { return false; }
		current = parent;
	}
	return true;
}

const problems = [];
let checked = 0;

for (const file of markdownFiles(ROOT)) {
	// Strip code before looking for links, or every example path becomes a false positive.
	const text = fs.readFileSync(file, 'utf8').replace(FENCE, '');
	const dir = path.dirname(file);

	let m;
	LINK.lastIndex = 0;
	while ((m = LINK.exec(text)) !== null) {
		const raw = m[1].trim().split(/\s+/)[0];
		if (!raw || isExternal(raw)) { continue; }

		// Drop any anchor; we check that the FILE exists, not the heading.
		const target = raw.split('#')[0];
		if (!target) { continue; }

		checked++;
		const resolved = path.resolve(dir, decodeURIComponent(target));
		const where = path.relative(ROOT, file).replace(/\\/g, '/');

		if (!fs.existsSync(resolved)) {
			problems.push({ file: where, target: raw, why: 'does not exist' });
		} else if (!existsWithExactCase(resolved)) {
			// Windows and macOS resolve this happily; GitHub and Linux do not. Without this check,
			// a link that works on every developer machine is broken for every reader.
			problems.push({ file: where, target: raw, why: 'wrong case' });
		}
	}
}

console.log(`Checked ${checked} relative links across the docs.`);

if (problems.length > 0) {
	console.error(`\n${problems.length} link(s) point at something that does not exist:\n`);
	for (const p of problems) {
		console.error(`  ${p.file}  ->  ${p.target}   (${p.why})`);
	}
	console.error('');
	process.exit(1);
}

console.log('All of them resolve.');
