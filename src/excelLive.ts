import { execFile } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Live sync: writing Power Query into a workbook Excel already has open.
 *
 * The extension's normal path rewrites the DataMashup part inside the .xlsx zip, and that fails
 * while Excel holds the file - Excel takes an exclusive WRITE lock. (Reading is fine, which is why
 * extraction never needed the file closed.)
 *
 * This module does not fight the lock. It asks the running Excel to make the change through its own
 * object model, so the file on disk is never touched, and the workbook is left dirty for the user
 * to save. It is a WRITE-PATH option, not a replacement for anything: extraction, bulk extraction
 * and writing to closed files all continue to work with no Excel, no COM and no Windows.
 *
 * WHY A SPAWNED SCRIPT rather than a native COM binding: a native module means prebuilt binaries
 * per Node ABI and a broken install for anyone whose ABI is not covered, in exchange for a feature
 * only Windows users can use at all. `powershell.exe` is already on every Windows machine.
 *
 * It must be powershell.exe specifically. `Marshal.GetActiveObject` does not exist in PowerShell 7
 * (.NET Core dropped it), and attaching to a RUNNING Excel is the entire point.
 */

/** How long to wait for the helper. Excel can be busy; it is not usually busy for ten seconds. */
const HELPER_TIMEOUT_MS = 10_000;

export interface LiveQuery {
	name: string;
	formula: string;
}

export interface LiveStatus {
	/** The helper ran and answered. False means fall back to the on-disk writer. */
	available: boolean;
	/** This particular workbook is open in the running Excel. */
	open: boolean;
	queries: string[];
	excelVersion?: string;
	workbook?: string;
	/** Machine-readable reason when `available` is false. */
	reason?: string;
	/**
	 * Excel processes running while this workbook was NOT found in the Running Object Table.
	 *
	 * A non-zero value with `open: false` is suspicious rather than normal - most often an
	 * integrity-level mismatch, which makes every workbook invisible without saying so.
	 */
	excelProcesses?: number;
	/**
	 * Excel reports the workbook has no unsaved changes.
	 *
	 * `false` means the user has edits in Excel that exist in NO file. The backup taken before a sync
	 * captures the LAST SAVED state, so it does not contain them - writing over those edits would
	 * destroy work that nothing anywhere has a copy of.
	 */
	saved?: boolean;
	/**
	 * Excel's AutoSave is on for this workbook - true for most OneDrive and SharePoint files.
	 *
	 * It changes what a live write means. Measured: with AutoSave on, a change written through COM
	 * is committed to disk within about two seconds, and closing the workbook without saving does
	 * NOT undo it. The "review before saving" behavior only exists when this is false.
	 *
	 * Absent rather than false when Excel would not report it.
	 */
	autoSaveOn?: boolean;
	/** Whether the helper itself ran elevated. Half of the mismatch above. */
	elevated?: boolean;
}

export interface LiveWriteResult {
	ok: boolean;
	open: boolean;
	updated: string[];
	added: string[];
	/** Present in the workbook already, byte-identical, so deliberately not written. */
	unchanged: string[];
	failures: { name: string; message: string }[];
	/** The workbook has unsaved changes - true after any real write. */
	dirty?: boolean;
	/** AutoSave is on, so Excel commits the write itself within seconds. See LiveStatus. */
	autoSaveOn?: boolean;
	reason?: string;
}

export function isLiveSyncSupported(): boolean {
	return process.platform === 'win32';
}

function helperPath(extensionPath: string): string {
	return path.join(extensionPath, 'resources', 'live-sync', 'excel-live-sync.ps1');
}

/**
 * Run the helper and parse its single line of JSON.
 *
 * The helper answers with JSON for failures too, so a rejected promise here means something more
 * basic went wrong - PowerShell missing, the script absent, a timeout.
 */
function runHelper(
	extensionPath: string,
	args: string[],
	stdin?: string
): Promise<Record<string, unknown>> {
	return new Promise((resolve, reject) => {
		const script = helperPath(extensionPath);
		if (!fs.existsSync(script)) {
			reject(new Error(`Live sync helper not found: ${script}`));
			return;
		}

		const child = execFile(
			'powershell.exe',
			['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args],
			{ timeout: HELPER_TIMEOUT_MS, windowsHide: true, maxBuffer: 4 * 1024 * 1024 },
			(error, stdout, stderr) => {
				if (error && !stdout) {
					reject(new Error(stderr?.trim() || error.message));
					return;
				}
				const line = stdout.trim().split(/\r?\n/).filter(Boolean).pop();
				if (!line) {
					reject(new Error('Live sync helper produced no output'));
					return;
				}
				try {
					resolve(JSON.parse(line) as Record<string, unknown>);
				} catch {
					reject(new Error(`Live sync helper returned unparseable output: ${line.slice(0, 200)}`));
				}
			}
		);

		// The payload is the user's source code. It goes down a pipe that exists only for the life
		// of this call - never a temp file, which would leave their M sitting on disk, and never a
		// command-line argument, where quotes and backslashes become a corruption bug.
		if (stdin !== undefined) {
			child.stdin?.on('error', () => { /* the callback above reports the real failure */ });
			child.stdin?.end(stdin, 'utf8');
		}
	});
}

/** Is this workbook open in Excel right now, and what queries does it hold? */
export async function getLiveStatus(workbookPath: string, extensionPath: string): Promise<LiveStatus> {
	if (!isLiveSyncSupported()) {
		return { available: false, open: false, queries: [], reason: 'not-windows' };
	}

	try {
		const r = await runHelper(extensionPath, ['-Action', 'status', '-Path', workbookPath]);
		if (r.ok === false) {
			// 'excel-not-running' is the ordinary case, not a fault.
			return { available: false, open: false, queries: [], reason: String(r.error ?? 'unknown') };
		}
		return {
			available: true,
			open: r.open === true,
			queries: Array.isArray(r.queries) ? (r.queries as string[]) : [],
			excelVersion: r.excelVersion as string | undefined,
			workbook: r.workbook as string | undefined,
			excelProcesses: typeof r.excelProcesses === 'number' ? r.excelProcesses : undefined,
			saved: typeof r.saved === 'boolean' ? r.saved : undefined,
			autoSaveOn: typeof r.autoSaveOn === 'boolean' ? r.autoSaveOn : undefined,
			elevated: typeof r.elevated === 'boolean' ? r.elevated : undefined
		};
	} catch (e) {
		return {
			available: false, open: false, queries: [],
			reason: e instanceof Error ? e.message : String(e)
		};
	}
}

/**
 * Write formulas into the open workbook.
 *
 * The payload goes down the helper's STDIN. Not a command-line argument, where M's quotes and
 * backslashes become a corruption bug, and not a temp file, which would write the user's source
 * code to disk behind their back.
 */
export async function writeLive(
	workbookPath: string,
	queries: LiveQuery[],
	extensionPath: string
): Promise<LiveWriteResult> {
	const empty: LiveWriteResult = { ok: false, open: false, updated: [], added: [], unchanged: [], failures: [] };

	// Nothing to write succeeds everywhere. There is nothing platform-specific about doing
	// nothing, and reporting 'not-windows' for an empty request made Linux CI fail on a call
	// that had asked for no work at all.
	if (queries.length === 0) {
		return { ok: true, open: true, updated: [], added: [], unchanged: [], failures: [] };
	}
	if (!isLiveSyncSupported()) {
		return { ...empty, reason: 'not-windows' };
	}

	try {
		const r = await runHelper(
			extensionPath,
			['-Action', 'write', '-Path', workbookPath],
			JSON.stringify(queries)
		);

		if (r.ok === false && r.error) {
			return { ...empty, open: r.open === true, reason: String(r.error) };
		}
		return {
			ok: r.ok === true,
			open: r.open === true,
			updated: Array.isArray(r.updated) ? (r.updated as string[]) : [],
			added: Array.isArray(r.added) ? (r.added as string[]) : [],
			unchanged: Array.isArray(r.unchanged) ? (r.unchanged as string[]) : [],
			failures: Array.isArray(r.failures)
				? (r.failures as { name: string; message: string }[])
				: [],
			dirty: r.dirty as boolean | undefined,
			autoSaveOn: typeof r.autoSaveOn === 'boolean' ? r.autoSaveOn : undefined
		};
	} catch (e) {
		return { ...empty, reason: e instanceof Error ? e.message : String(e) };
	}
}

/**
 * A human-readable explanation for "Excel is clearly running, but we cannot see the workbook".
 *
 * Returns undefined when there is nothing suspicious to explain.
 */
export function explainInvisibleWorkbook(status: LiveStatus): string | undefined {
	if (status.open || !status.excelProcesses) { return undefined; }
	return 'Excel is running but this workbook is not visible to the extension. '
		+ 'This usually means one of them is elevated and the other is not - '
		+ 'COM hides running objects across integrity levels. '
		+ 'Run VS Code and Excel the same way (normally, for preference) and try again.';
}

/**
 * Should a live sync refuse because the workbook holds unsaved work?
 *
 * Extracted so the decision can be tested directly. Verifying it end to end means driving a real
 * Excel into a dirty state, which is exactly the situation we refuse to gamble with.
 *
 * `saved === undefined` means the helper did not report - an older helper, or a status call that
 * could not reach the workbook. Do NOT refuse on unknown: that would block every sync the moment the
 * field went missing. Refuse only on a definite `false`.
 */
export function shouldRefuseUnsavedWorkbook(status: Pick<LiveStatus, 'open' | 'saved'>): boolean {
	return status.open === true && status.saved === false;
}

/**
 * Can this file be written right now?
 *
 * Opens for read/write WITHOUT truncating and writes nothing, then closes immediately. If Windows
 * refuses (EBUSY/EPERM), something holds an exclusive lock - which is the only question worth asking
 * first. A workbook that is not locked is not open, so there is nothing for live sync to do and the
 * ordinary file writer handles it without Excel being consulted at all.
 *
 * Verified non-destructive: the file's mtime does not move.
 */
export function isWritableNow(filePath: string): boolean {
	try {
		const fd = fs.openSync(filePath, 'r+');
		fs.closeSync(fd);
		return true;
	} catch {
		return false;
	}
}

/** Excel leaves `~$Name.xlsx` beside a workbook it has open FROM DISK. */
export function hasOwnerLockFile(filePath: string): boolean {
	const sibling = path.join(path.dirname(filePath), '~$' + path.basename(filePath));
	return fs.existsSync(sibling);
}

/**
 * The file cannot be written AND no Excel we can see has it open.
 *
 * This is a real, ordinary situation rather than a bug: COM partitions the running object table by
 * integrity level, so a workbook open in an elevated Excel is invisible to a normal-integrity
 * extension, and vice versa. We can prove the file is locked - the write fails - but we cannot reach
 * whatever holds it, and we must not guess at a similarly-named workbook we CAN see. Measured on a
 * real machine, that guess wrote into a different workbook two directories away.
 *
 * Both ways out are the user's to choose, so say both.
 */
/**
 * Why live sync did not handle a locked workbook - and it is usually not what we used to say.
 *
 * REPORTED BY A USER, AND THEY WERE RIGHT. The old code showed one message whenever the workbook
 * was locked and live sync had not handled it, and that message blamed integrity levels. But the
 * probe is gated on `sync.liveWhenOpen`, which is OFF by default - so the overwhelmingly common
 * case was "you have not turned it on", answered with a paragraph about COM and elevation.
 *
 * Their log said `owner lock file present: true`, meaning Excel had the workbook open FROM DISK as
 * that same user. We had the evidence to rule elevation out and showed the elevation message
 * anyway. Enabling the setting fixed it, after a round trip that should never have been needed.
 *
 * Kept pure and separated from the UI so every branch is testable without a workbook, an Excel, or
 * a running COM server.
 */
export function explainLiveSyncUnavailable(opts: {
	fileName: string;
	/** `sync.liveWhenOpen` - off by default. */
	enabled: boolean;
	/** Windows, in practice. */
	supported: boolean;
	/** `~$Name.xlsx` beside the workbook: Excel opened it from disk, as this user. */
	ownerLockPresent: boolean;
}): { message: string; offerEnable: boolean } {
	const { fileName, enabled, supported, ownerLockPresent } = opts;

	if (!supported) {
		return {
			message: `${fileName} is open in Excel, so it cannot be written directly. Live sync can `
				+ 'write into an open workbook, but it needs Windows and Excel - it works by asking '
				+ 'Excel to make the change. Close the workbook and sync again.',
			offerEnable: false
		};
	}

	if (!enabled) {
		return {
			message: `${fileName} is open in Excel. This extension can write straight into an open `
				+ 'workbook, but that is off by default - turn on '
				+ '`excel-power-query-editor.sync.liveWhenOpen` and sync again. Or close the '
				+ 'workbook and sync as usual.',
			offerEnable: true
		};
	}

	if (ownerLockPresent) {
		// Excel holds it, from disk, as this user - so elevation is the wrong thing to blame.
		return {
			message: `${fileName} is open in Excel and live sync is on, but the workbook could not `
				+ 'be reached through Excel. Set `excel-power-query-editor.log.level` to `debug` and '
				+ 'check the Output panel for the `live sync` line - it names the reason. Closing the '
				+ 'workbook and syncing again works meanwhile.',
			offerEnable: false
		};
	}

	return { message: explainLockedButUnreachable(fileName), offerEnable: false };
}

export function explainLockedButUnreachable(filePath: string): string {
	return `${path.basename(filePath)} is open somewhere this extension cannot reach. The usual `
		+ 'cause is Excel and VS Code running at different privilege levels - COM hides running '
		+ 'objects across integrity levels, so we can tell the file is locked but cannot talk to '
		+ 'whatever holds it. '
		+ 'Either run VS Code and Excel the same way, or close the workbook and sync again - with it '
		+ 'closed the file is written directly and none of this applies.';
}
