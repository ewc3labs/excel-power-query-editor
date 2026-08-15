import * as assert from 'assert';
import * as vscode from 'vscode';
import { migrateLegacySettings } from '../../src/extension';

/**
 * The settings migration, exercised against WORKSPACE scope.
 *
 * WHY THIS FILE EXISTS SEPARATELY. The default test host runs with no folder open, so
 * `ConfigurationTarget.Workspace` cannot be written at all and every existing migration test asserts
 * on Global. That is precisely how the bug below survived a full suite, a CI matrix, and a review.
 *
 * THE BUG. The migration marker is written to Global scope, but the guard read the EFFECTIVE value
 * of `xtn.level` - global merged over workspace and folder. So once ANY workspace had been migrated,
 * the guard reported "already done" in every other workspace, forever. Nothing reads the old keys as
 * a fallback, so a skipped workspace does not quietly keep working: those settings stop applying and
 * the defaults take over. Someone with `backup.location` set to `tempFolder` in one project starts
 * writing backups next to a workbook in their synced OneDrive folder and is never told.
 *
 * The reproduction is the middle test: mark Global as migrated, put a legacy value in WORKSPACE
 * scope, and require that it still migrates.
 */

const SECTION = 'excel-power-query-editor';
const MARKER = 'xtn.level';
const SCHEMA = '1';

const G = vscode.ConfigurationTarget.Global;
const W = vscode.ConfigurationTarget.Workspace;

function config(): vscode.WorkspaceConfiguration {
	return vscode.workspace.getConfiguration(SECTION);
}

async function set(key: string, value: unknown, target: vscode.ConfigurationTarget): Promise<void> {
	await config().update(key, value, target);
}

function scoped(key: string, field: 'globalValue' | 'workspaceValue'): unknown {
	const info = config().inspect(key);
	return info ? (info as Record<string, unknown>)[field] : undefined;
}

suite('Settings migration - workspace scope', () => {
	const touched = [
		MARKER, 'watchAlways', 'watch.always', 'backupLocation', 'backup.location',
		'syncTimeout', 'sync.timeout',
	];

	async function clearAll(): Promise<void> {
		for (const key of touched) {
			await set(key, undefined, G);
			await set(key, undefined, W);
		}
	}

	setup(async () => { await clearAll(); });
	teardown(async () => { await clearAll(); });

	test('a folder really is open, or nothing below proves anything', () => {
		assert.ok(
			(vscode.workspace.workspaceFolders?.length ?? 0) > 0,
			'this suite must run in a host with a workspace folder; check .vscode-test.mjs'
		);
	});

	test('migrates a legacy value held in workspace scope', async () => {
		await set('watchAlways', true, W);

		await migrateLegacySettings();

		assert.strictEqual(scoped('watch.always', 'workspaceValue'), true,
			'the workspace value should have been carried across');
		assert.strictEqual(scoped('watchAlways', 'workspaceValue'), undefined,
			'the old workspace key should have been cleared');
	});

	test('a Global marker must NOT skip workspace scope', async () => {
		// The regression. This is the state a user is in the moment any one workspace has migrated:
		// the marker is set globally, and a different project still holds legacy settings of its own.
		await set(MARKER, SCHEMA, G);
		await set('backupLocation', 'tempFolder', W);

		await migrateLegacySettings();

		assert.strictEqual(scoped('backup.location', 'workspaceValue'), 'tempFolder',
			'a workspace with its own legacy settings must still migrate, even though Global is marked');
		assert.strictEqual(scoped('backupLocation', 'workspaceValue'), undefined,
			'the old workspace key should have been cleared');
	});

	test('does not overwrite a new workspace value the user already set', async () => {
		await set('syncTimeout', 10, W);
		await set('sync.timeout', 99, W);

		await migrateLegacySettings();

		assert.strictEqual(scoped('sync.timeout', 'workspaceValue'), 99,
			'an explicit new value wins over a legacy one');
		assert.strictEqual(scoped('syncTimeout', 'workspaceValue'), undefined,
			'the stale old key is still cleared');
	});

	test('marks the workspace only when something actually moved', async () => {
		// A clean workspace must not get our bookkeeping written into its committed
		// .vscode/settings.json. Rescanning it costs a few inspect() calls; a spurious marker costs
		// somebody a line in their next git diff and an explanation they did not ask for.
		await migrateLegacySettings();

		assert.strictEqual(scoped(MARKER, 'workspaceValue'), undefined,
			'nothing moved here, so nothing should have been written to workspace scope');
	});
});
