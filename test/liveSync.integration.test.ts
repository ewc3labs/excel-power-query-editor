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

/** A multi-line script, passed as -EncodedCommand so no quoting or newline survives the command line wrong. */
function psScript(script: string): string {
	return execFileSync('powershell.exe',
		['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')],
		{ encoding: 'utf8', timeout: 60000 }).trim();
}

/** A literal for a single-quoted PowerShell string: only ' is special, and it is escaped by doubling. */
function psQuote(value: string): string {
	return `'${value.replace(/'/g, "''")}'`;
}

suite('Live sync against a real Excel', function () {
	// Excel is slow to start and slower to quit.
	this.timeout(180_000);

	const available = excelInstalled();
	// Only an Excel this suite started may be quit by this suite - identified by PROCESS ID, not by
	// "whichever Excel COM hands us", which would be someone else's the moment they opened one.
	let excelPid = 0;
	const workdir = path.join(os.tmpdir(), 'epqe-live-integration');
	const workbook = path.join(workdir, 'live-integration.xlsx');
	let extensionPath = '';

	suiteSetup(function () {
		if (!available) {
			console.log('    [skipped] Excel is not available on this machine');
			this.skip();
			return;
		}

		// REFUSE TO RUN IF EXCEL IS ALREADY OPEN.
		//
		// This suite drives a real Excel and has to clean up after itself. There is no way to do that
		// safely alongside a developer's own session: COM hands out the one running instance, so any
		// cleanup we perform lands on THEIR workbooks. An earlier version of this teardown closed
		// every open workbook with SaveChanges=$false and quit Excel - which, on any machine with
		// Excel installed, is exactly the silent data loss this extension exists to avoid.
		//
		// So: if Excel is already running, skip. A skipped test costs a line of output. The other
		// outcome costs someone their morning.
		const running = Number(ps('@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count')) || 0;
		if (running > 0) {
			console.log('    [skipped] Excel is already running - close it to run the live sync suite');
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

		// Open it the way a user does, then wait for Excel to register it in the ROT. -PassThru returns
		// the EXCEL process itself (measured 2026-09-14), which is how teardown knows what is ours.
		excelPid = Number(ps(`(Start-Process ${psQuote(workbook)} -PassThru).Id`)) || 0;
		if (!excelPid) { throw new Error('Start-Process did not report the Excel process it launched'); }
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
		if (!available || !excelPid) { return; }

		// THE SUITE USED TO DISABLE ITSELF AFTER ONE RUN. Two defects, found 2026-09-14:
		//
		// 1. The workbook path was put in a single-quoted PowerShell string with every backslash
		//    doubled. Backslash is not an escape there, so the path never equalled FullName, nothing
		//    closed, and Excel was never quit.
		// 2. Even closed and quit correctly, EXCEL DOES NOT EXIT. Quit() is accepted and the process
		//    stays, with zero workbooks and no window - reproduced on every trial: quit at 0s, 45s and
		//    3 minutes after launch, with Saved set and COM references released, in /safe mode with
		//    no add-ins, and after a COM formula write. Same PID throughout. Cause not found.
		//
		// Either way the next run saw "Excel is already running" and skipped all five tests.
		//
		// So identify our Excel by the PID we launched, and if it outlives Quit() with NOTHING open,
		// end that process. If anything is open in it - someone opened a file mid-run and it joined
		// our instance - leave it and say so. Ending a process with unsaved work in it is the one
		// outcome this suite exists to never cause.
		const cs = path.join(extensionPath, 'resources', 'live-sync', 'RunningObjects.cs.txt');
		try {
			const outcome = psScript(`
				$excelPid = ${excelPid}
				Add-Type -TypeDefinition (Get-Content -Raw ${psQuote(cs)})
				$book = [RunningObjects]::Get(${psQuote(workbook)})
				if ($null -ne $book) {
					$app = $book.Application
					$book.Close($false)
					if (@($app.Workbooks).Count -eq 0) { $app.Quit() }
					$book = $null; $app = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
				}
				$deadline = (Get-Date).AddSeconds(15)
				while ((Get-Date) -lt $deadline -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 500 }
				if (-not (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { 'exited'; return }
				$all = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
				if ($all.Count -ne 1) { "left running: $($all.Count) Excel processes, so COM may not be pointing at ours"; return }
				$open = @([Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application').Workbooks).Count
				if ($open -eq 0) { Stop-Process -Id $excelPid -Force; 'ended: outlived Quit with nothing open' }
				else { "left running: $open workbook(s) open in it" }
			`);
			console.log(`    [teardown] Excel ${excelPid}: ${outcome.split(/\r?\n/).pop()}`);
		} catch (e) {
			// Not silent. A swallowed teardown failure is how this suite disabled itself unnoticed.
			console.log(`    [teardown] Excel ${excelPid}: teardown failed - ${e instanceof Error ? e.message : e}`);
		}
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
