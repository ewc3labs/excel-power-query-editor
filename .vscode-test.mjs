import { defineConfig } from '@vscode/test-cli';

export default defineConfig({
	files: 'out/test/**/*.test.js',
	mocha: {
		// Mocha's default is 2000ms, which is the wrong budget for this suite: these are integration
		// tests that open real .xlsx workbooks, write files, and round-trip settings through VS Code's
		// configuration API. On a developer machine that fits easily; on a loaded CI runner it does
		// not, and the result is a timeout on a different test each run - a flake that looks like a
		// bug in whatever happened to be slowest that day.
		//
		// Raising this does not hide a hang. A genuinely stuck test still fails, twenty seconds later.
		timeout: 20000
	}
});
