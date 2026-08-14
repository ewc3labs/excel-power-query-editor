import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execFileSync } from 'child_process';
import * as vscode from 'vscode';
import { getLiveStatus, writeLive, isLiveSyncSupported } from '../src/excelLive';
import { parseSection, buildSection, detectEol } from '../src/mSection';

/**
 * The acid test for the two-tailed write: drive a real Excel, through the real helper.
 *
 * SKIPS ITSELF unless Windows has Excel installed, because CI has neither. That is a deliberate
 * trade - this cannot be the only place live sync is verified, which is why the unit tests cover
 * the contract when Excel is absent. What only this test can prove is that a formula written
 * through the object model comes back out of the FILE identical to what the section document said.
 */

function excelInstalled(): boolean {
	if (process.platform !== 'win32') { return false; }
	try {
		const out = execFileSync('powershell.exe',
			['-NoProfile', '-NonInteractive', '-Command',
				'try { $null = [Type]::GetTypeFromProgID("Excel.Application"); if ($null -ne [Type]::GetTypeFromProgID("Excel.Application")) { "yes" } else { "no" } } catch { "no" }'],
			{ encoding: 'utf8', timeout: 15000 }).trim();
		return out.endsWith('yes');
	} catch {
		return false;
	}
}

function ps(command: string): string {
	return execFileSync('powershell.exe',
		['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', command],
		{ encoding: 'utf8', timeout: 60000 }).trim();
}

suite('Live sync against a real Excel', function () {
	// Excel is slow to start and slower to quit.
	this.timeout(180_000);

	const available = excelInstalled();
	const workdir = path.join(os.tmpdir(), 'epqe-live-integration');
	const workbook = path.join(workdir, 'live-integration.xlsx');
	let extensionPath = '';

	suiteSetup(function () {
		if (!available) {
			console.log('    [skipped] Excel is not available on this machine');
			this.skip();
			return;
		}
		extensionPath = vscode.extensions
			.getExtension('ewc3labs.excel-power-query-editor')?.extensionPath
			?? path.join(__dirname, '..', '..');

		fs.mkdirSync(workdir, { recursive: true });
		fs.copyFileSync(
			path.join(__dirname, '..', '..', 'test', 'fixtures', 'simple.xlsx'),
			workbook
		);

		// Open it the way a user does, then wait for Excel to register it in the ROT.
		ps(`Start-Process '${workbook}'`);
		const deadline = Date.now() + 90_000;
		while (Date.now() < deadline) {
			const listed = ps(
				`Add-Type -TypeDefinition (Get-Content -Raw '${path.join(extensionPath, 'resources', 'live-sync', 'RunningObjects.cs.txt')}');` +
				`if ([RunningObjects]::Names() -contains '${workbook}') { 'ready' } else { 'waiting' }`
			);
			if (listed.endsWith('ready')) { return; }
			execFileSync('powershell.exe', ['-NoProfile', '-Command', 'Start-Sleep -Seconds 2']);
		}
		throw new Error('Excel never registered the workbook in the ROT');
	});

	suiteTeardown(() => {
		if (!available) { return; }
		try {
			ps('try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); ' +
				'foreach ($w in @($xl.Workbooks)) { $w.Close($false) }; $xl.Quit() } catch { }');
		} catch { /* best effort */ }
		try { fs.rmSync(workdir, { recursive: true, force: true }); } catch { /* best effort */ }
	});

	test('the workbook is seen as open, with its queries', async () => {
		const status = await getLiveStatus(workbook, extensionPath);
		assert.strictEqual(status.available, true, 'helper should be able to answer');
		assert.strictEqual(status.open, true, 'the workbook we just opened is open');
		assert.ok(status.queries.includes('StudentResults'), `queries were ${status.queries.join(',')}`);
	});

	test('a closed workbook next to it is NOT reported open, and is NOT opened', async () => {
		const closed = path.join(workdir, 'closed-neighbour.xlsx');
		fs.copyFileSync(workbook, closed);

		const before = ps('@(Get-Process EXCEL -EA SilentlyContinue).Count');
		const status = await getLiveStatus(closed, extensionPath);
		const after = ps('@(Get-Process EXCEL -EA SilentlyContinue).Count');

		assert.strictEqual(status.available, true, 'we could ask');
		assert.strictEqual(status.open, false, 'and the answer is no');
		assert.strictEqual(before, after, 'checking must not start or change Excel');
	});

	test('a live write reaches the workbook and leaves the FILE alone', async () => {
		const before = fs.statSync(workbook).mtimeMs;
		const marker = `// live integration ${Date.now()}`;
		const status = await getLiveStatus(workbook, extensionPath);
		const original = status.queries;

		const result = await writeLive(
			workbook,
			[{ name: 'StudentResults', formula: `let\r\n    ${marker}\r\n    Source = 1\r\nin\r\n    Source` }],
			extensionPath
		);

		assert.strictEqual(result.ok, true, JSON.stringify(result));
		assert.deepStrictEqual(result.updated, ['StudentResults']);
		assert.strictEqual(result.dirty, true, 'the workbook should now have unsaved changes');
		assert.strictEqual(fs.statSync(workbook).mtimeMs, before,
			'the file on disk must be untouched - that is the whole point');
		assert.deepStrictEqual((await getLiveStatus(workbook, extensionPath)).queries, original,
			'writing a formula must not change which queries exist');
	});

	test('PQ-15: a section document survives the round trip through Excel', async () => {
		// section document -> split -> N formulas -> Excel -> read back -> section document
		const source = fs.readFileSync(
			path.join(__dirname, '..', '..', 'test', 'fixtures', 'expected', 'simple_StudentResults.m'),
			'utf8');
		const parsed = parseSection(source);

		const written = await writeLive(
			workbook,
			parsed.queries.map(q => ({ name: q.name, formula: q.expression })),
			extensionPath
		);
		assert.strictEqual(written.ok, true, JSON.stringify(written));

		// Read the expression back out of Excel and rebuild the document around it.
		const readBack = ps(
			`Add-Type -TypeDefinition (Get-Content -Raw '${path.join(extensionPath, 'resources', 'live-sync', 'RunningObjects.cs.txt')}');` +
			`$wb = [RunningObjects]::Get('${workbook}');` +
			`$wb.Queries.Item('StudentResults').Formula`
		);

		const rebuilt = buildSection(
			{ header: parsed.header, queries: [{ ...parsed.queries[0], expression: readBack.replace(/\r?\n/g, '\r\n') }] },
			detectEol(source)
		);

		assert.strictEqual(rebuilt, source,
			'a document written through Excel and read back must be byte-identical');
	});

	test('live sync is only offered where it can work', () => {
		assert.strictEqual(isLiveSyncSupported(), true, 'this suite only runs on Windows');
	});
});
