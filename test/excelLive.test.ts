import * as assert from 'assert';
import * as path from 'path';
import * as os from 'os';
import * as fs from 'fs';
import { execFile } from 'child_process';
import { isLiveSyncSupported, getLiveStatus, writeLive, shouldRefuseUnsavedWorkbook, explainLiveSyncUnavailable, explainInvisibleWorkbook } from '../src/excelLive';

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

suite('Live sync payload encoding', function () {
	this.timeout(30_000);

	const helper = path.join(__dirname, '..', '..', 'resources', 'live-sync', 'excel-live-sync.ps1');

	/*
	 * THE BUG THIS GUARDS. Node writes the M payload to the helper as UTF-8. PowerShell's
	 * [Console]::In decodes using [Console]::InputEncoding, which falls back to the system OEM
	 * codepage when no console is attached - and the extension spawns the helper with windowsHide,
	 * so there never is one.
	 *
	 * An em-dash is E2 80 94 in UTF-8. Those three bytes decoded as CP437 are the string a user
	 * reported seeing in round-tripped M. Every non-ASCII character had the same problem: accented
	 * names, smart quotes, degree signs.
	 *
	 * It survived manual testing because an interactive shell is usually already at 65001. The bug
	 * only existed under the spawn conditions the extension actually uses.
	 */

	test('the helper reads stdin with an explicit UTF-8 encoding', () => {
		// Runs everywhere, including CI on Linux, and it is the assertion that fails if somebody
		// reverts to the terser form. A true end-to-end check would need Excel and an open workbook,
		// which no CI runner has - so this guards the fix rather than re-proving the mechanism.
		const src = fs.readFileSync(helper, 'utf8');

		assert.ok(
			/OpenStandardInput\(\)[\s\S]{0,400}UTF8Encoding/.test(src),
			'stdin must be read through a StreamReader with an explicit UTF8Encoding'
		);
		assert.ok(
			!/\[Console\]::In\.ReadToEnd\(\)/.test(src),
			'[Console]::In.ReadToEnd() decodes with the console codepage - that was the bug'
		);
		assert.ok(
			/\[Console\]::OutputEncoding\s*=/.test(src),
			'the return trip needs pinning too, or non-ASCII mangles on the way back'
		);
	});

	test('non-ASCII survives the stdin construct the helper uses', async function () {
		if (process.platform !== 'win32') { this.skip(); return; }

		// Behavioural, and deliberately narrow: it runs the SAME reader construct through a real
		// powershell.exe and asserts the bytes come back intact. It cannot exercise the helper end to
		// end without Excel, so it proves the mechanism rather than the whole path.
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'epqe-enc-'));
		const payloadFile = path.join(dir, 'payload.txt');
		const readerFile = path.join(dir, 'read.ps1');

		const emDash = '—';
		fs.writeFileSync(payloadFile, `BCMDB2 (Baylor) ${emDash} doctor NPIs`, 'utf8');
		fs.writeFileSync(readerFile, [
			'$stdin  = [Console]::OpenStandardInput()',
			'$reader = New-Object System.IO.StreamReader($stdin, (New-Object System.Text.UTF8Encoding $false))',
			'$raw    = $reader.ReadToEnd()',
			'[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
			'$b = [System.Text.UTF8Encoding]::new($false).GetBytes($raw)',
			"Write-Output (($b | ForEach-Object { $_.ToString('X2') }) -join ' ')",
		].join(String.fromCharCode(10)), 'utf8');

		const hex: string = await new Promise((resolve, reject) => {
			const child = execFile(
				'powershell.exe',
				['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', readerFile],
				{ windowsHide: true },
				(err, stdout) => (err && !stdout ? reject(err) : resolve(stdout.trim()))
			);
			child.stdin?.end(fs.readFileSync(payloadFile));
		});

		assert.ok(
			hex.includes('E2 80 94'),
			`the em-dash did not survive as UTF-8. Got: ${hex.slice(0, 120)}`
		);
	});
});

suite('Why live sync did not run', function () {
	/*
	 * REPORTED BY A USER. The locked-workbook message used to blame integrity levels in every case,
	 * but the probe is gated on sync.liveWhenOpen - off by default - so the common case was "not
	 * turned on", answered with a paragraph about COM elevation. Their log even said the owner lock
	 * file was present, which rules elevation out.
	 */
	const base = { fileName: 'Book.xlsx', enabled: true, supported: true, ownerLockPresent: false };

	test('the setting being off says so, and offers to turn it on', () => {
		const r = explainLiveSyncUnavailable({ ...base, enabled: false });
		assert.ok(r.message.includes('sync.liveWhenOpen'), 'must name the setting to change');
		assert.strictEqual(r.offerEnable, true, 'the fix is one click, so offer it');
		assert.ok(!/elevat|privilege|integrity/i.test(r.message), 'must not blame elevation');
	});

	test('an unsupported platform is not a settings problem', () => {
		const r = explainLiveSyncUnavailable({ ...base, enabled: false, supported: false });
		assert.ok(/Windows/.test(r.message));
		assert.strictEqual(r.offerEnable, false, 'enabling it would not help');
	});

	test('an owner lock file rules elevation out, so it points at the log instead', () => {
		// ~$Book.xlsx means Excel opened it from disk as this user. Blaming integrity levels here is
		// contradicted by evidence we already have in hand.
		const r = explainLiveSyncUnavailable({ ...base, ownerLockPresent: true });
		assert.ok(!/elevat|privilege|integrity/i.test(r.message), 'the lock file disproves that');
		assert.ok(/debug|Output/i.test(r.message), 'send them where the real reason is printed');
		assert.strictEqual(r.offerEnable, false, 'it is already on');
	});

	test('enabled, supported, no lock file - elevation is a fair guess again', () => {
		const r = explainLiveSyncUnavailable(base);
		assert.ok(/privilege|elevat|integrity/i.test(r.message));
		assert.strictEqual(r.offerEnable, false);
	});
});

suite('Why a running Excel cannot see the workbook', function () {
	/*
	 * REPORTED FROM A NETWORK SHARE. The old message said "this usually means one of them is
	 * elevated" every time. The user's helper had measured itself NOT elevated, and the workbook was
	 * sitting in the Running Object Table under its UNC path. So these branch on what the helper SAW.
	 */
	const running = { available: true, open: false, queries: [], excelProcesses: 1 };

	test('nothing to explain when the workbook was found, or Excel is not running', () => {
		assert.strictEqual(explainInvisibleWorkbook({ ...running, open: true }), undefined);
		assert.strictEqual(explainInvisibleWorkbook({ ...running, excelProcesses: 0 }), undefined);
	});

	test('an elevated helper blames VS Code, specifically', () => {
		const m = explainInvisibleWorkbook({ ...running, elevated: true, registered: [] })!;
		assert.ok(/administrator/i.test(m));
		assert.ok(/VS Code is running as administrator/.test(m), 'we measured which side is elevated - say so');
	});

	test('zero visible workbooks while Excel runs is an integrity wall', () => {
		const m = explainInvisibleWorkbook({ ...running, elevated: false, registered: [] })!;
		assert.ok(/none of its workbooks/i.test(m));
		assert.ok(/administrator/i.test(m));
	});

	test('some workbooks visible rules elevation out and points at a different path', () => {
		// The network-drive report: we could see into Excel fine, the name just did not match.
		const m = explainInvisibleWorkbook({
			...running, elevated: false,
			registered: ['\\\\medarms01\\public\\IT\\Other.xlsx', 'C:\\Users\\me\\Book.xlsx']
		})!;
		assert.ok(!/elevat|administrator|integrity/i.test(m), 'we can see into Excel, so elevation is disproved');
		assert.ok(/2 workbooks are visible/.test(m), 'say how much we could see');
		assert.ok(/different path/i.test(m));
		assert.ok(/log/i.test(m), 'the names are in the log');
	});

	test('no evidence - an older helper, or a table the helper could not read - does not claim what is usual', () => {
		// `registered` absent is also what the helper sends when it FAILED to read the table. That must
		// not be mistaken for an empty table, which would blame an integrity wall nobody measured.
		const m = explainInvisibleWorkbook({ ...running, elevated: false })!;
		assert.ok(!/usually/i.test(m), 'we have no evidence either way, so do not rank the causes');
		assert.ok(/different path/i.test(m) && /administrator/i.test(m), 'name both possibilities');
	});
});

suite('Mapped network drives', function () {
	this.timeout(30_000);
	const helperDir = path.join(__dirname, '..', '..', 'resources', 'live-sync');

	test('the helper tries the UNC form after the exact path, before cloud URLs', () => {
		const src = fs.readFileSync(path.join(helperDir, 'excel-live-sync.ps1'), 'utf8');
		const exact = src.indexOf("how = 'exact-path'");
		const unc = src.indexOf('[NetworkPaths]::ToUnc(');
		const cloud = src.indexOf('ConvertTo-CloudUrl -LocalPath $FullPath');
		assert.ok(exact > 0 && unc > exact && cloud > unc, 'order must be exact, then UNC, then cloud');
	});

	test('the shipped C# still compiles, and ToUnc declines everything that is not a mapped drive', async function () {
		if (process.platform !== 'win32') { this.skip(); return; }

		// A compile error in RunningObjects.cs.txt would break EVERY live sync, not just network
		// ones, so compile the real file. A mapped drive cannot be created in CI without changing the
		// machine, and ToUnc's POSITIVE case has not run anywhere yet. What was verified by hand on
		// 2026-09-14 is only that Excel registered a P:\ workbook under \\medarms01\public in the ROT.
		const cs = path.join(helperDir, 'RunningObjects.cs.txt').replace(/'/g, "''");
		const script = [
			`Add-Type -TypeDefinition (Get-Content -Raw '${cs}') -ErrorAction Stop`,
			"$r = @('C:\\Windows\\win.ini', '\\\\server\\share\\x.xlsx', '', 'relative\\x.xlsx') | ForEach-Object { [string][NetworkPaths]::ToUnc($_) }",
			"Write-Output ('RESULT:' + ($r -join '|'))",
			// Names() returns NULL only when the table cannot be read. On a working machine it must be a
			// real array, even an empty one - otherwise null would mean something besides failure.
			"$n = [RunningObjects]::Names(); Write-Output ('NAMES:' + ($null -ne $n) + ':' + ($n -is [array]))",
		].join('; ');

		const out: string = await new Promise((resolve, reject) => {
			execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { windowsHide: true },
				(err, stdout, stderr) => (err ? reject(new Error(stderr || String(err))) : resolve(stdout)));
		});
		const line = out.split(/\r?\n/).find((l) => l.startsWith('RESULT:'));
		assert.ok(line, `helper C# did not compile or run: ${out}`);
		assert.strictEqual(line, 'RESULT:|||', 'local, UNC, empty and relative paths all have no second name');
		const names = out.split(/\r?\n/).find((l) => l.startsWith('NAMES:'));
		assert.strictEqual(names, 'NAMES:True:True', 'a readable table yields an array, so null is reserved for failure');
	});
});
