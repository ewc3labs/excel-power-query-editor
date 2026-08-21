import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';

/**
 * Give the Power Query language service the Excel symbols it does not ship.
 *
 * WHY THIS IS NEEDED AT ALL. Measured against `powerquery.vscode-powerquery` 1.0.0: its bundled
 * standard library contains `Excel.Workbook` - read *an* external workbook file - and has zero
 * occurrences of `Excel.CurrentWorkbook`, which reads *this* workbook and is where nearly every
 * in-workbook query begins. Without it, the first line of almost every real file lints as an
 * unknown identifier.
 *
 * WHY THIS REPLACED WRITING A FILE. The previous approach wrote `excel-pq-symbols.json` into
 * `<workspace>/.vscode/` and appended that folder to the user's
 * `powerquery.client.additionalSymbolsDirectories` setting. Three things wrong with it:
 *
 *   - it threw outright when there was no workspace, which is exactly how somebody tries this
 *     extension for the first time: open a single .m file and see what happens. It has been
 *     failing on every CI run for a year for the same reason.
 *   - `.vscode/` is version-controlled in most projects, so it created a file in a repository the
 *     user never asked us to touch, to be committed and puzzled over by their colleagues.
 *   - it edited another extension's settings, permanently, to solve a problem that lasts only as
 *     long as the session.
 *
 * The Power Query extension exposes an API for precisely this - `addLibrarySymbols`, which the PQ
 * SDK itself uses to push in connector symbols. Nothing is written anywhere, no setting is touched,
 * and the symbols disappear with the session that added them. See
 * https://github.com/microsoft/vscode-powerquery/issues/206.
 */

const POWER_QUERY_EXTENSION = 'powerquery.vscode-powerquery';

/** Our entry in the language service's symbol table. Also the key used to remove it again. */
const LIBRARY_ID = 'excel-power-query-editor';

/**
 * The subset of the Power Query extension's API we use.
 *
 * Declared here rather than imported: taking a dependency on their package to describe three
 * methods would be heavier than the methods themselves, and the extension may not be installed at
 * all - in which case there is nothing to type against.
 *
 * >>> UPSTREAM CONTRACT: microsoft/vscode-powerquery#206 <<<
 * https://github.com/microsoft/vscode-powerquery/issues/206
 *
 * THESE METHOD NAMES ARE NOT A PUBLISHED API. They come from that issue, and TypeScript cannot
 * check them for us because the shape is declared here rather than imported. If Power Query renames
 * or removes them, nothing fails to compile - the `typeof` guard below simply starts returning
 * `api-not-available` forever and Excel symbols quietly stop appearing.
 *
 * Grep `vscode-powerquery#206` to find every place that assumption is made.
 */
interface PowerQueryApi {
	readonly addLibrarySymbols: (
		librarySymbols: ReadonlyMap<string, ReadonlyArray<unknown>>
	) => Promise<void>;
	readonly removeLibrarySymbols: (librariesToRemove: ReadonlyArray<string>) => Promise<void>;
}

export interface SymbolRegistrationResult {
	ok: boolean;
	/** How many symbols were handed over. */
	count: number;
	/** Why it did not happen, when it did not. */
	reason?: string;
}

function symbolsFile(extensionPath: string): string {
	return path.join(extensionPath, 'resources', 'symbols', 'excel-pq-symbols.json');
}

/**
 * Hand our symbols to the Power Query extension.
 *
 * Safe to call when the PQ extension is absent - that is simply a user who has not installed it,
 * and it is not our place to complain about that on their behalf.
 */
export async function registerExcelSymbols(
	extensionPath: string,
	log: (message: string, level?: string) => void
): Promise<SymbolRegistrationResult> {
	const pq = vscode.extensions.getExtension<PowerQueryApi>(POWER_QUERY_EXTENSION);
	if (!pq) {
		return { ok: false, count: 0, reason: 'power-query-extension-not-installed' };
	}

	let api: PowerQueryApi;
	try {
		// It may not have started yet; activation is idempotent and returns the same exports.
		api = pq.isActive ? pq.exports : await pq.activate();
	} catch (e) {
		return { ok: false, count: 0, reason: `activation-failed: ${e instanceof Error ? e.message : e}` };
	}

	if (!api || typeof api.addLibrarySymbols !== 'function') {
		// An older Power Query extension predating the API. Nothing to do, and nothing broken.
		//
		// THIS IS ALSO WHERE A RENAME LANDS. `addLibrarySymbols` is from vscode-powerquery#206 and
		// is not a published contract, so an upstream rename is indistinguishable from an old
		// version here - both are "the method is not there". If Excel symbols stop appearing for
		// users on a CURRENT Power Query extension, check the method name upstream before anything
		// else. https://github.com/microsoft/vscode-powerquery/issues/206
		return { ok: false, count: 0, reason: 'api-not-available' };
	}

	const file = symbolsFile(extensionPath);
	let symbols: ReadonlyArray<unknown>;
	try {
		symbols = JSON.parse(fs.readFileSync(file, 'utf8'));
	} catch (e) {
		return { ok: false, count: 0, reason: `symbols-unreadable: ${e instanceof Error ? e.message : e}` };
	}

	if (!Array.isArray(symbols) || symbols.length === 0) {
		return { ok: false, count: 0, reason: 'symbols-empty' };
	}

	try {
		await api.addLibrarySymbols(new Map([[LIBRARY_ID, symbols]]));
		log(`Registered ${symbols.length} Excel symbols with the Power Query language service`, 'info');
		return { ok: true, count: symbols.length };
	} catch (e) {
		return { ok: false, count: 0, reason: `add-failed: ${e instanceof Error ? e.message : e}` };
	}
}

/**
 * Take them back out again.
 *
 * Called on deactivate so we leave the language service as we found it. Failures are deliberately
 * silent: the extension is shutting down, and there is nobody left to tell.
 */
export async function unregisterExcelSymbols(): Promise<void> {
	try {
		const pq = vscode.extensions.getExtension<PowerQueryApi>(POWER_QUERY_EXTENSION);
		if (pq?.isActive && typeof pq.exports?.removeLibrarySymbols === 'function') {
			await pq.exports.removeLibrarySymbols([LIBRARY_ID]);
		}
	} catch {
		// Nothing useful to do while shutting down.
	}
}

/** A sentence a human can act on, for each way this can decline. */
export function explainRegistration(result: SymbolRegistrationResult): string {
	if (result.ok) {
		return `Excel symbols are active — ${result.count} registered with the Power Query language service.`;
	}
	switch (result.reason) {
		case 'power-query-extension-not-installed':
			return 'Install the Power Query / M Language extension to get IntelliSense for '
				+ 'Excel.CurrentWorkbook().';
		case 'api-not-available':
			// Says "too old" because that is the likely cause - but an upstream RENAME of
			// `addLibrarySymbols` presents identically and would make this sentence wrong for
			// everybody. See vscode-powerquery#206 and the guard in registerExcelSymbols.
			return 'The installed Power Query extension is too old to accept extra symbols. '
				+ 'Update it to get Excel.CurrentWorkbook() IntelliSense.';
		default:
			return `Could not register Excel symbols: ${result.reason}`;
	}
}

/**
 * Keep the symbols registered when the Power Query extension comes and goes.
 *
 * Registering once at activation is not enough. VS Code loads extensions in an order nobody
 * controls, and a user can install, enable, disable or update the Power Query extension long after
 * we started. In every one of those cases our symbols are either never delivered or quietly
 * dropped, and the failure is invisible: `Excel.CurrentWorkbook()` simply goes back to being an
 * unknown identifier and the user has no reason to connect that to an extension they just touched.
 *
 * `extensions.onDidChange` fires for install, uninstall, enable and disable. Re-registering is
 * cheap and idempotent - the same library id replaces the previous entry - so the safe move is to
 * try again whenever the set of extensions changes at all.
 */
export function watchForPowerQueryExtension(
	extensionPath: string,
	log: (message: string, level?: string) => void
): vscode.Disposable {
	let lastKnownPresent = vscode.extensions.getExtension(POWER_QUERY_EXTENSION) !== undefined;

	return vscode.extensions.onDidChange(async () => {
		const presentNow = vscode.extensions.getExtension(POWER_QUERY_EXTENSION) !== undefined;

		// It went away. Nothing to do - it took our symbols with it, and it will get them back if
		// it returns.
		if (!presentNow) {
			if (lastKnownPresent) {
				log('Power Query extension is gone; Excel symbols went with it', 'debug');
			}
			lastKnownPresent = false;
			return;
		}

		lastKnownPresent = true;

		const result = await registerExcelSymbols(extensionPath, log);
		if (result.ok) {
			log('Power Query extension changed; Excel symbols re-registered', 'info');
		} else if (result.reason !== 'power-query-extension-not-installed') {
			log(`Power Query extension changed but symbols did not register: ${result.reason}`, 'debug');
		}
	});
}

/**
 * LEFTOVERS FROM THE FILE-BASED VERSION, WHICH WE FIND AND REPORT BUT NEVER DELETE.
 *
 * Before PQ-18 this extension wrote `excel-pq-symbols/excel-pq-symbols.json` to disk and appended
 * that folder to `powerquery.client.additionalSymbolsDirectories`. Upgrading does not remove either,
 * because an upgrade has no business deleting files a user might have edited, moved, or come to
 * depend on - the same reason live sync refuses to overwrite an unsaved workbook.
 *
 * But leaving them silent is its own problem: if the setting still points at the old folder, the
 * Power Query extension keeps loading those symbols from disk while we push the current ones through
 * the API. At best that is redundant, and the stale copy never updates again.
 *
 * So: find them, say so once, and let the human decide.
 */

const LEGACY_DIR_NAME = 'excel-pq-symbols';
const LEGACY_FILE_NAME = 'excel-pq-symbols.json';

export interface LegacyLeftovers {
	/** Directories that still contain our old symbols file. */
	readonly folders: ReadonlyArray<string>;
	/** True when `additionalSymbolsDirectories` still names one of our folders. */
	readonly settingStillPoints: boolean;
}

/** Both places the old version could have written, matching its own path logic exactly. */
function legacyCandidates(): string[] {
	const out: string[] = [];

	// User scope. The old code used APPDATA or HOME - not a VS Code API - so we repeat that rather
	// than compute a "better" path and miss the files it actually wrote.
	const userDataPath = process.env.APPDATA || process.env.HOME;
	if (userDataPath) {
		out.push(path.join(userDataPath, 'Code', 'User', LEGACY_DIR_NAME));
	}

	for (const folder of vscode.workspace.workspaceFolders ?? []) {
		out.push(path.join(folder.uri.fsPath, '.vscode', LEGACY_DIR_NAME));
	}

	return out;
}

/** Look for what the file-based version left behind. Reads only; writes and deletes nothing. */
export function findLegacyLeftovers(): LegacyLeftovers {
	const folders = legacyCandidates().filter(dir => {
		try {
			return fs.existsSync(path.join(dir, LEGACY_FILE_NAME));
		} catch {
			// An unreadable path is not our problem to report.
			return false;
		}
	});

	let settingStillPoints = false;
	try {
		const configured = vscode.workspace
			.getConfiguration('powerquery.client')
			.get<string[]>('additionalSymbolsDirectories') ?? [];
		settingStillPoints = configured.some(d =>
			typeof d === 'string' && d.replace(/[\/]+$/, '').endsWith(LEGACY_DIR_NAME));
	} catch {
		// The Power Query extension may not be installed, so the section may not exist.
	}

	return { folders, settingStillPoints };
}
