import * as assert from 'assert';
import * as path from 'path';
import * as os from 'os';
import * as fs from 'fs';
import { isLiveSyncSupported, getLiveStatus, writeLive, shouldRefuseUnsavedWorkbook } from '../src/excelLive';

/**
 * Live sync needs Windows, Excel, and a workbook actually open - none of which CI has. So these
 * cover the contract that matters when those things are ABSENT: every path returns a value the
 * caller can fall back on, and nothing throws.
 *
 * The COM behavior itself is verified by hand against a running Excel; see
 * docs/design/live-sync-to-open-excel.md for what was measured.
 */
suite('Live sync', function () {
	// Not the 2s default. Every call here spawns powershell.exe, which runs Add-Type and COMPILES
	// C# at runtime - a few hundred milliseconds on a warm machine and several seconds on a cold CI
	// runner. These passed locally and failed in CI purely on that, with one of the passing ones
	// taking 1842ms of the 2000ms allowed.
	this.timeout(30_000);

	const extensionPath = path.join(__dirname, '..', '..');

	test('support is platform-gated, and honestly reported', () => {
		assert.strictEqual(isLiveSyncSupported(), process.platform === 'win32');
	});

	test('status never throws - it reports why it cannot help', async () => {
		const status = await getLiveStatus(
			path.join(os.tmpdir(), 'definitely-not-open-' + Date.now() + '.xlsx'),
			extensionPath
		);
		assert.strictEqual(typeof status.available, 'boolean');
		assert.strictEqual(status.open, false, 'a file nobody has open is not open');
		assert.ok(Array.isArray(status.queries));
		if (!status.available) {
			assert.ok(status.reason, 'unavailable must come with a reason the caller can log');
		}
	});

	test('a missing helper degrades instead of exploding', async () => {
		const status = await getLiveStatus(
			path.join(os.tmpdir(), 'anything.xlsx'),
			path.join(os.tmpdir(), 'no-extension-here-' + Date.now())
		);
		assert.strictEqual(status.available, false);
		assert.ok(status.reason, 'the reason should name the missing helper');
	});

	test('writing nothing is a success, not a round trip to Excel', async () => {
		const result = await writeLive(
			path.join(os.tmpdir(), 'whatever.xlsx'), [], extensionPath
		);
		assert.strictEqual(result.ok, true);
		assert.deepStrictEqual(result.updated, []);
		assert.deepStrictEqual(result.added, []);
	});

	test('a write to a workbook nobody has open reports open=false, and never throws', async () => {
		const result = await writeLive(
			path.join(os.tmpdir(), 'not-open-' + Date.now() + '.xlsx'),
			[{ name: 'Q', formula: 'let x = 1 in x' }],
			extensionPath
		);
		assert.strictEqual(result.open, false, 'caller uses this to choose the on-disk writer');
		assert.ok(Array.isArray(result.failures));
	});

	test('the payload never touches disk', async () => {
		const before = fs.readdirSync(os.tmpdir()).filter(f => f.startsWith('epqe-live-')).length;
		await writeLive(
			path.join(os.tmpdir(), 'not-open.xlsx'),
			[{ name: 'Q', formula: 'let secret = "user M code" in secret' }],
			extensionPath
		);
		const after = fs.readdirSync(os.tmpdir()).filter(f => f.startsWith('epqe-live-')).length;
		// It is the user's source code. It travels on stdin and is never written anywhere.
		assert.strictEqual(after, before, 'no payload file should ever be created');
	});

	test('the shipped helper is where the code looks for it', () => {
		const helper = path.join(extensionPath, 'resources', 'live-sync', 'excel-live-sync.ps1');
		assert.ok(fs.existsSync(helper), `helper must ship at ${helper}`);
		const body = fs.readFileSync(helper, 'utf8');
		// Not GetActiveObject: that finds ONE instance and misses a workbook open in another.
		assert.ok(body.includes('[RunningObjects]::Get'),
			'helper must find the workbook through the ROT, not a single Excel instance');
		assert.ok(!body.includes('New-Object -ComObject Excel.Application'),
			'helper must never START Excel - that would surprise the user');
		// The CALL syntax, not the word: the helper's own comments explain why BindToMoniker is
		// avoided, and matching those would make this assertion pass for the wrong reason.
		assert.ok(!body.includes(']::BindToMoniker'),
			'BindToMoniker OPENS a closed file - measured. The ROT lookup exists to avoid that.');
		assert.ok(fs.existsSync(path.join(extensionPath, 'resources', 'live-sync', 'RunningObjects.cs.txt')),
			'the ROT interop source must ship alongside the helper');
	});
});

suite('Refusing to overwrite unsaved work', () => {
	// The backup taken before a sync copies the file on DISK. If Excel holds edits the user has not
	// saved, those edits are in no file at all, so the backup looks like protection and is not.
	test('refuses when the workbook is open with unsaved changes', () => {
		assert.strictEqual(shouldRefuseUnsavedWorkbook({ open: true, saved: false }), true);
	});

	test('allows a clean open workbook', () => {
		assert.strictEqual(shouldRefuseUnsavedWorkbook({ open: true, saved: true }), false);
	});

	test('does NOT refuse when the helper did not report', () => {
		// An older helper, or a status call that could not reach the workbook. Refusing on unknown
		// would block every sync the moment this field went missing.
		assert.strictEqual(shouldRefuseUnsavedWorkbook({ open: true, saved: undefined }), false);
	});

	test('is irrelevant when the workbook is not open', () => {
		assert.strictEqual(shouldRefuseUnsavedWorkbook({ open: false, saved: false }), false);
	});
});
