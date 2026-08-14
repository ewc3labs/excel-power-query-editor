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
			workbook: r.workbook as string | undefined
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

	if (!isLiveSyncSupported()) {
		return { ...empty, reason: 'not-windows' };
	}
	if (queries.length === 0) {
		return { ok: true, open: true, updated: [], added: [], unchanged: [], failures: [] };
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
			dirty: r.dirty as boolean | undefined
		};
	} catch (e) {
		return { ...empty, reason: e instanceof Error ? e.message : String(e) };
	}
}
