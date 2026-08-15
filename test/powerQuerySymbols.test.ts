import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import {
	registerExcelSymbols,
	unregisterExcelSymbols,
	explainRegistration,
	watchForPowerQueryExtension
} from '../src/powerQuerySymbols';

/**
 * The Power Query extension is a SEPARATE extension that we do not depend on and cannot require.
 * It may be absent, present, older than the API, or installed halfway through a session.
 *
 * The test host runs with only this extension loaded, so `powerquery.vscode-powerquery` is
 * genuinely absent here - which makes this suite a real test of the not-installed path rather than
 * a simulated one. When run in a development host that DOES have it, the same tests assert the
 * successful path instead. Both are checked below rather than assumed.
 */
suite('Excel symbols', function () {
	this.timeout(20_000);

	const extensionPath = path.join(__dirname, '..', '..');
	const powerQueryPresent = vscode.extensions.getExtension('powerquery.vscode-powerquery') !== undefined;
	const noop = () => { /* the tests assert on the result, not the log */ };

	suiteSetup(() => {
		console.log(`    [Power Query extension ${powerQueryPresent ? 'IS' : 'is NOT'} installed here]`);
	});

	test('the symbols we ship are well formed and match what the language service expects', () => {
		const file = path.join(extensionPath, 'resources', 'symbols', 'excel-pq-symbols.json');
		assert.ok(fs.existsSync(file), 'the symbols file must ship');

		const symbols = JSON.parse(fs.readFileSync(file, 'utf8'));
		assert.ok(Array.isArray(symbols) && symbols.length > 0, 'must be a non-empty array');

		// The shape was checked against the entry the language server holds for Excel.Workbook.
		// completionItemKind and type - NOT completionItemType/dataType, which an older revision
		// of the API declaration suggested.
		for (const s of symbols) {
			assert.ok(typeof s.name === 'string' && s.name.length > 0, 'every symbol needs a name');
			assert.ok('completionItemKind' in s, `${s.name} needs completionItemKind`);
			assert.ok('type' in s, `${s.name} needs type`);
			assert.ok('documentation' in s, `${s.name} needs documentation`);
		}

		// The reason this file exists at all.
		assert.ok(symbols.some((s: { name: string }) => s.name === 'Excel.CurrentWorkbook'),
			'Excel.CurrentWorkbook is the symbol the language service does not ship');
	});

	test('registering never throws, whether or not Power Query is installed', async () => {
		const result = await registerExcelSymbols(extensionPath, noop);
		assert.strictEqual(typeof result.ok, 'boolean');

		if (!powerQueryPresent) {
			assert.strictEqual(result.ok, false, 'cannot register without the extension');
			assert.strictEqual(result.reason, 'power-query-extension-not-installed');
		} else if (result.ok) {
			assert.ok(result.count > 0, 'a successful registration hands over symbols');
		}
	});

	test('a missing Power Query extension is explained, not reported as a fault', () => {
		const message = explainRegistration({
			ok: false, count: 0, reason: 'power-query-extension-not-installed'
		});
		assert.ok(/install the power query/i.test(message),
			'the message should tell the user what to do, not what failed');
		assert.ok(!/error|failed/i.test(message),
			'not having an optional extension is not an error');
	});

	test('an outdated Power Query extension is explained too', () => {
		const message = explainRegistration({ ok: false, count: 0, reason: 'api-not-available' });
		assert.ok(/update/i.test(message), 'tell them the actionable thing');
	});

	test('unregistering is safe when there is nothing to unregister', async () => {
		// Runs on deactivate, possibly with the Power Query extension already gone.
		await unregisterExcelSymbols();
		await unregisterExcelSymbols();
	});

	test('the watcher can be created and disposed without a Power Query extension', () => {
		const watcher = watchForPowerQueryExtension(extensionPath, noop);
		assert.ok(watcher, 'a disposable is returned');
		watcher.dispose();
		// Disposing twice is what happens if activation fails halfway.
		watcher.dispose();
	});

	test('a bad extension path is reported, not thrown', async () => {
		const result = await registerExcelSymbols(
			path.join(extensionPath, 'no-such-directory-' + Date.now()), noop
		);
		assert.strictEqual(result.ok, false);
		assert.ok(result.reason, 'a reason the caller can log');
		if (powerQueryPresent) {
			assert.ok(/symbols-unreadable/.test(result.reason!),
				'with the extension present, the failure should be about the missing file');
		}
	});
});
