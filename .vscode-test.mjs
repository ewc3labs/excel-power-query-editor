import { defineConfig } from '@vscode/test-cli';

// Mocha's default is 2000ms, which is the wrong budget for this suite: these are integration tests
// that open real .xlsx workbooks, write files, and round-trip settings through VS Code's
// configuration API. On a developer machine that fits easily; on a loaded CI runner it does not, and
// the result is a timeout on a different test each run - a flake that looks like a bug in whatever
// happened to be slowest that day.
//
// Raising this does not hide a hang. A genuinely stuck test still fails, twenty seconds later.
const mocha = { timeout: 20000 };

// macOS caps a Unix domain socket path at 104 bytes, and VS Code puts its IPC socket inside the
// user-data directory. GitHub Actions checks out to work/<repo>/<repo>, so the default
// .vscode-test/user-data/... path came to 106 characters and the host died with
// `listen EINVAL: invalid argument`. Two bytes over.
//
// Somewhere short, on macOS only - Windows and Linux have no such limit and are already green.
const launchArgs = process.platform === 'darwin' ? ['--user-data-dir=/tmp/epqe-vsc'] : [];

export default defineConfig([
	{
		label: 'unit',
		files: ['out/test/*.test.js'],
		launchArgs,
		mocha
	},
	{
		// A SECOND HOST, WITH A FOLDER OPEN.
		//
		// The default host runs with no workspace folder, so ConfigurationTarget.Workspace and
		// WorkspaceFolder cannot be written at all - which is exactly why a settings migration bug
		// that only affected those scopes went unnoticed through a full test suite and a review.
		// Anything asserting on workspace-scoped configuration has to live here.
		label: 'workspace',
		files: ['out/test/workspace/*.test.js'],
		workspaceFolder: './test/fixtures/migration-workspace',
		launchArgs,
		mocha
	}
]);
