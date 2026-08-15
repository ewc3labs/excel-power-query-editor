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
