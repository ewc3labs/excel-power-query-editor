// Command to open extension settings UI filtered to this extension
function openExtensionSettings() {
	vscode.commands.executeCommand('workbench.action.openSettings', '@excel-power-query-editor');
}
// The module 'vscode' contains the VS Code extensibility API
import * as vscode from 'vscode';
import { parseSection, diffQueries } from './mSection';
import { registerExcelSymbols, unregisterExcelSymbols, explainRegistration, watchForPowerQueryExtension, findLegacyLeftovers } from './powerQuerySymbols';
import {
	explainInvisibleWorkbook,
	explainLockedButUnreachable,
	getLiveStatus,
	hasOwnerLockFile,
	isLiveSyncSupported,
	shouldRefuseUnsavedWorkbook,
	writeLive
} from './excelLive';
import * as fs from 'fs';
import * as path from 'path';
import { watch, FSWatcher } from 'chokidar';
import { getConfig } from './configHelper';

/**
 * Are we running under a test host?
 *
 * The `describe` probe is deliberately written through an index signature rather than as
 * `global.describe`. The latter only compiles because @types/mocha declares `describe` globally, so
 * production source was silently depending on TEST types being in scope - which type-checks from the
 * command line and errors in an editor resolving a narrower context. Asking a plain object for a
 * property it may not have is the honest expression of what this is doing.
 */
function isTestEnvironment(): boolean {
	// Note what this probe does NOT include: mocha's TDD interface defines `suite`, not `describe`,
	// so under this project's own test host it is `undefined` and detection falls to the environment
	// variables. Adding a `suite` probe looks like an obvious completion of the thought and is not -
	// it flips this true during the suite, activating test-only branches that were never active, and
	// nineteen tests change behavior. The probe stays exactly as narrow as it has always been.
	const g = globalThis as unknown as Record<string, unknown>;
	return process.env.NODE_ENV === 'test'
		|| process.env.VSCODE_TEST_ENV === 'true'
		|| typeof g.describe !== 'undefined';
}

// Helper to get test fixture path  
function getTestFixturePath(filename: string): string {
	return path.join(__dirname, '..', 'test', 'fixtures', filename);
}

// File watchers storage 
const fileWatchers = new Map<string, { chokidar: FSWatcher; vscode: vscode.FileSystemWatcher | null; document: vscode.Disposable | null }>();
const recentExtractions = new Set<string>(); // Track recently extracted files to prevent immediate auto-sync

// Debounce timers for file sync operations
const debounceTimers = new Map<string, NodeJS.Timeout>();

// Output channel for verbose logging
let outputChannel: vscode.LogOutputChannel;

// Status bar item for watch status
let statusBarItem: vscode.StatusBarItem;

// Log level constants (external so they're not recreated every call)
const LOG_LEVEL_PRIORITY: { [key: string]: number } = {
	'none': 0, 'debug': 1, 'verbose': 2, 'info': 3, 'success': 3, 'warn': 4, 'error': 5
};

const LOG_LEVEL_EMOJIS: { [key: string]: string } = {
	'debug': '🪲',      // bug
	'verbose': '🔍',    // magnifying glass
	'info': 'ℹ️',       // info icon
	'success': '✅',    // checkmark
	'warn': '⚠️',       // warning triangle
	'error': '❌',      // X mark
	'none': '🚫'        // prohibition
};

const LOG_LEVEL_LABELS: { [key: string]: string } = {
	'debug': '[DEBUG]',
	'verbose': '[VERBOSE]',
	'info': '[INFO]',
	'success': '[SUCCESS]',
	'warn': '[WARN]',
	'error': '[ERROR]',
	'none': '[NONE]'
};

function supportsEmoji(): boolean {
	// VS Code output panel always supports emoji
	// Check if we're running in VS Code environment
	if (typeof vscode !== 'undefined') {
		return true;
	}
	
	// Fallback for other environments
	const platform = process.platform;
	// Modern terminals generally support emojis
	return platform !== 'win32' || !!process.env.TERM_PROGRAM || !!process.env.WT_SESSION;
}

// Backup path helper
function getBackupPath(excelFile: string, timestamp: string): string {
	const config = getConfig();
	const backupLocation = config.get<string>('backup.location', 'sameFolder');
	const baseFileName = path.basename(excelFile);
	const backupFileName = `${baseFileName}.backup.${timestamp}`;
	
	switch (backupLocation) {
		case 'tempFolder':
			return path.join(require('os').tmpdir(), 'excel-pq-backups', backupFileName);
		case 'custom':
			const customPath = config.get<string>('backup.customPath', '');
			if (customPath) {
				// Resolve relative paths relative to the Excel file directory
				const resolvedPath = path.isAbsolute(customPath) 
					? customPath 
					: path.resolve(path.dirname(excelFile), customPath);
				return path.join(resolvedPath, backupFileName);
			}
			// Fall back to same folder if custom path is not set
			return path.join(path.dirname(excelFile), backupFileName);
		case 'sameFolder':
		default:
			return path.join(path.dirname(excelFile), backupFileName);
	}
}

// Backup cleanup helper
function cleanupOldBackups(excelFile: string): void {
	const config = getConfig();
	const maxBackups = config.get<number>('backup.maxFiles', 5) || 5;
	const autoCleanup = config.get<boolean>('backup.autoCleanup', true) || false;
	
	if (!autoCleanup || maxBackups <= 0) {
		return;
	}
	
	try {
		// Get the backup directory based on settings
		const sampleTimestamp = '2000-01-01T00-00-00-000Z';
		const sampleBackupPath = getBackupPath(excelFile, sampleTimestamp);
		const backupDir = path.dirname(sampleBackupPath);
		const baseFileName = path.basename(excelFile);
		
		if (!fs.existsSync(backupDir)) {
			return;
		}
		
		// Find all backup files for this Excel file
		const backupPattern = `${baseFileName}.backup.`;
		const allFiles = fs.readdirSync(backupDir);
		const backupFiles = allFiles
			.filter(file => file.startsWith(backupPattern))
			.map(file => {
				const fullPath = path.join(backupDir, file);
				const timestampMatch = file.match(/\.backup\.(.+)$/);
				const timestamp = timestampMatch ? timestampMatch[1] : '';
				return {
					path: fullPath,
					filename: file,
					timestamp: timestamp,
					// Parse timestamp for sorting (ISO format sorts naturally)
					sortKey: timestamp
				};
			})
			.filter(backup => backup.timestamp) // Only files with valid timestamps
			.sort((a, b) => b.sortKey.localeCompare(a.sortKey)); // Newest first
		
		// Delete excess backups
		if (backupFiles.length > maxBackups) {
			const filesToDelete = backupFiles.slice(maxBackups);
			let deletedCount = 0;
			
			for (const backup of filesToDelete) {
				try {
					fs.unlinkSync(backup.path);
					deletedCount++;
					log(`Deleted old backup: ${backup.filename}`, 'cleanupOldBackups', 'debug');
				} catch (deleteError) {
					log(`Failed to delete backup ${backup.filename}: ${deleteError}`, 'cleanupOldBackups', 'error');
				}
			}
			
			if (deletedCount > 0) {
				log(`Cleaned up ${deletedCount} old backup files (keeping ${maxBackups} most recent)`, 'cleanupOldBackups', 'info');
			}
		}
		
	} catch (error) {
		log(`Backup cleanup failed: ${error}`, 'cleanupOldBackups', 'error');
	}
}

// Enhanced logging function with context and log levels, smart emoji or text 'level' support, and respects user log level settings
function log(message: string, context: string = '', level: string = 'info'): void {
	const config = getConfig();
	const userLogLevel = (config.get<string>('log.level', 'info') || 'info').toLowerCase();
	const messageLevel = level.toLowerCase();

	const userPriority = LOG_LEVEL_PRIORITY[userLogLevel] ?? 3;
	const messagePriority = LOG_LEVEL_PRIORITY[messageLevel] ?? 3;

	// If user set 'none', suppress all logging, or if message is below threshold
	if (userLogLevel === 'none' || messagePriority < userPriority) {
		return;
	}

	const timestamp = new Date().toISOString();
	const emojiMode = supportsEmoji();
	const levelSymbol = emojiMode
		? LOG_LEVEL_EMOJIS[messageLevel] || 'ℹ️'
		: LOG_LEVEL_LABELS[messageLevel] || '[INFO]';

	let logPrefix = `[${timestamp}] ${levelSymbol}`;
	if (context) {
		logPrefix += ` [${context}]`;
	}

	const fullMessage = `${logPrefix} ${message}`;
	console.log(fullMessage);
	
	if (outputChannel) {
		// Level-specific methods, so VS Code can filter and persist correctly. `message`
		// rather than `fullMessage`: the channel supplies its own timestamp and level.
		const line = context ? `[${context}] ${message}` : message;
		switch (messageLevel) {
			case 'error': outputChannel.error(line); break;
			case 'warn': outputChannel.warn(line); break;
			case 'debug': outputChannel.debug(line); break;
			case 'verbose': outputChannel.trace(line); break;
			default: outputChannel.info(line); break;
		}
	}
}

// Update status bar
function updateStatusBar() {
	const config = getConfig();
	if (!config.get<boolean>('log.showStatusBarInfo', true)) {
		statusBarItem?.hide();
		return;
	}

	if (!statusBarItem) {
		statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
	}

	const watchedFiles = fileWatchers.size;
	if (watchedFiles > 0) {
		statusBarItem.text = `$(eye) Watching ${watchedFiles} PQ file${watchedFiles > 1 ? 's' : ''}`;
		statusBarItem.tooltip = `Power Query files being watched: ${Array.from(fileWatchers.keys()).map(f => path.basename(f)).join(', ')}`;
		statusBarItem.show();
	} else {
		statusBarItem.hide();
	}
}

// Initialize auto-watch for existing .m files
async function initializeAutoWatch(): Promise<void> {
	const config = getConfig();
	const watchAlways = config.get<boolean>('watch.always', false);
	
	if (!watchAlways) {
		log('Extension activated - auto-watch disabled, staying dormant until manual command', 'initializeAutoWatch', 'info');
		return; // Auto-watch is disabled - minimal initialization
	}

	log('Extension activated - auto-watch enabled, scanning workspace for .m files...', 'initializeAutoWatch', 'info');

	try {
		// Find all .m files in the workspace
		const mFiles = await vscode.workspace.findFiles('**/*.m', '**/node_modules/**');
		
		if (mFiles.length === 0) {
			log('Auto-watch enabled but no .m files found in workspace', 'initializeAutoWatch', 'info');
			vscode.window.showInformationMessage('🔍 Auto-watch enabled but no .m files found in workspace');
			return;
		}

		log(`Found ${mFiles.length} .m files in workspace, checking for corresponding Excel files...`, 'initializeAutoWatch', 'verbose');

		let watchedCount = 0;
		const maxAutoWatch = config.get<number>('watch.maxFiles', 25) || 25; // Configurable limit for auto-watch
		
		if (mFiles.length > maxAutoWatch) {
			log(`Found ${mFiles.length} .m files but limiting auto-watch to ${maxAutoWatch} files (configurable in settings)`, 'initializeAutoWatch', 'info');
		}
		
		for (const mFileUri of mFiles.slice(0, maxAutoWatch)) {
			const mFile = mFileUri.fsPath;
			
			// Check if there's a corresponding Excel file
			const excelFile = await findExcelFile(mFile);
			if (excelFile && fs.existsSync(excelFile)) {
				try {
					await watchFile(mFileUri);
					watchedCount++;
					log(`Auto-watch initialized: ${path.basename(mFile)} → ${path.basename(excelFile)}`, 'initializeAutoWatch', 'debug');
				} catch (error) {
					log(`Failed to auto-watch ${path.basename(mFile)}: ${error}`, 'initializeAutoWatch', 'error');
				}
			} else {
				log(`Skipping ${path.basename(mFile)} - no corresponding Excel file found`, 'initializeAutoWatch', 'debug');
			}
		}

		if (watchedCount > 0) {
			vscode.window.showInformationMessage(
				`🚀 Auto-watch enabled: Now watching ${watchedCount} Power Query file${watchedCount > 1 ? 's' : ''}`
			);
			log(`Auto-watch initialization complete: ${watchedCount} files being watched`, 'initializeAutoWatch', 'info');
		} else {
			log('Auto-watch enabled but no .m files with corresponding Excel files found', 'initializeAutoWatch', 'info');
			vscode.window.showInformationMessage('⚠️ Auto-watch enabled but no .m files with corresponding Excel files found');
		}

		if (mFiles.length > maxAutoWatch) {
			vscode.window.showWarningMessage(
				`Found ${mFiles.length} .m files but only auto-watching first ${maxAutoWatch}. Use "Watch File" command for others.`
			);
			log(`Limited auto-watch to ${maxAutoWatch} files (found ${mFiles.length} total)`, 'initializeAutoWatch', 'warn');
		}

	} catch (error) {
		log(`Auto-watch initialization failed: ${error}`, 'initializeAutoWatch', 'error');
		vscode.window.showErrorMessage(`Auto-watch initialization failed: ${error}`);
	}
}

// This method is called when your extension is activated
export async function activate(context: vscode.ExtensionContext) {
	try {
		// Initialize output channel first (before any logging)
		// `{ log: true }` makes this a LogOutputChannel - what VS Code expects an extension to
		// use. It carries a level per message, honors the user's own log-level control
		// ("Developer: Set Log Level..."), and VS Code PERSISTS IT TO DISK itself, under
		// the window's exthost log folder. That is what makes a bug report legible.
		//
		// An earlier attempt here wrote a log file by hand into %LOCALAPPDATA%. That is a
		// desktop-application pattern: extensions do not own folders in a user's profile,
		// and it meant reimplementing rolling and cleanup VS Code already does. If a real
		// file is ever needed, `context.logUri` is the sanctioned place.
		outputChannel = vscode.window.createOutputChannel('Excel Power Query Editor', { log: true });
		
		const self = vscode.extensions.getExtension('ewc3labs.excel-power-query-editor');
		log(`Excel Power Query Editor active - version ${self?.packageJSON?.version ?? 'unknown'}, `
			+ `host ${vscode.env.appName}, extensionPath ${self?.extensionPath ?? 'unknown'}`,
			'activate', 'info');

		// Register all commands
		// Migrate legacy settings (debugMode/verboseMode) to logLevel
		await migrateLegacySettings();
		const commands = [
			vscode.commands.registerCommand('excel-power-query-editor.extractFromExcel', extractFromExcel),
			vscode.commands.registerCommand('excel-power-query-editor.syncToExcel', syncToExcel),
			vscode.commands.registerCommand('excel-power-query-editor.watchFile', watchFile),
			vscode.commands.registerCommand('excel-power-query-editor.toggleWatch', toggleWatch),
			vscode.commands.registerCommand('excel-power-query-editor.stopWatching', stopWatching),
			vscode.commands.registerCommand('excel-power-query-editor.syncAndDelete', syncAndDelete),
			vscode.commands.registerCommand('excel-power-query-editor.rawExtraction', rawExtraction),
			vscode.commands.registerCommand('excel-power-query-editor.cleanupBackups', cleanupBackupsCommand),
			vscode.commands.registerCommand('excel-power-query-editor.installExcelSymbols', installExcelSymbols),
			vscode.commands.registerCommand('excel-power-query-editor.openSettings', openExtensionSettings)
		];

		context.subscriptions.push(...commands);
		log(`Registered ${commands.length} commands successfully`, 'activate', 'success');

		// Initialize status bar
		updateStatusBar();

		log('Excel Power Query Editor extension activated', 'activate', 'info');
		
		// Auto-watch existing .m files if setting is enabled
		await initializeAutoWatch();
		
		// Hand our Excel symbols to the Power Query language service. No file is written and no
		// setting is changed, so there is nothing to gate behind a preference and nothing to
		// fail when no workspace is open - which is what the old file-based version did on
		// every CI run for a year.
		// Symbols go to the POWER QUERY extension's API, not to any VS Code API - see
		// vscode-powerquery#206. If it is absent, old, or renames the method, registerExcelSymbols
		// declines with a reason and nothing here breaks.
		//
		// NOT AWAITED, DELIBERATELY. registerExcelSymbols calls `activate()` on the Power Query
		// extension, so awaiting it makes OUR activation wait on THEIRS - and the only thing we do
		// with the result is write one debug line. The try/catch in there covers Power Query
		// throwing; it cannot cover Power Query being slow, and activation time is the one number
		// VS Code shows users in `Extensions: Show Running Extensions`.
		//
		// Nothing downstream needs the symbols to be registered before we finish activating, and
		// watchForPowerQueryExtension below re-registers on any later arrival, so a late finish
		// costs nothing.
		void registerExcelSymbols(context.extensionPath,
			(m, l) => log(m, 'excelSymbols', l ?? 'info'))
			.then(symbolResult => {
				if (!symbolResult.ok) {
					log(`Excel symbols not registered: ${symbolResult.reason}`, 'excelSymbols', 'debug');
				}
			});

		// And keep them registered. The Power Query extension can be installed, enabled or updated
		// after we start, and in each case symbols registered at activation are never delivered or
		// are silently dropped.
		context.subscriptions.push(
			watchForPowerQueryExtension(context.extensionPath,
				(m, l) => log(m, 'excelSymbols', l ?? 'info'))
		);
		
		// Tell the user about anything the OLD file-based symbols version left on disk. Once.
		//
		// We do not delete it. An upgrade deleting a user's files is a trade nobody agreed to, and
		// they may have edited or moved that copy. Reporting is the whole job; the decision is
		// theirs. See vscode-powerquery#206 and PQ-18 for why the file stopped being used.
		void reportLegacySymbolLeftovers(context);

		log('Extension activation completed successfully', 'activate', 'success');
	} catch (error) {
		log(`Extension activation failed: ${error}`, 'activate', 'error');
		// Re-throw to ensure VS Code knows about the failure
		throw error;
	}
}

async function extractFromExcel(uri?: vscode.Uri, uris?: vscode.Uri[]): Promise<void> {
	try {
		// Dump extension settings for debugging (debug level only)
		const logLevel = getConfig().get<string>('log.level', 'info');
		if (logLevel === 'debug') {
			dumpAllExtensionSettings();
		}
		
		// Handle multiple file selection (batch operations)
		if (uris && uris.length > 1) {
			log(`Batch extraction started: ${uris.length} files selected`, 'extractFromExcel', 'info');
			vscode.window.showInformationMessage(`Extracting Power Query from ${uris.length} Excel files...`);
			
			let successCount = 0;
			let errorCount = 0;
			
			for (const fileUri of uris) {
				try {
					await extractFromExcel(fileUri); // Recursive call for single file
					successCount++;
				} catch (error) {
					log(`Failed to extract from ${path.basename(fileUri.fsPath)}: ${error}`, 'extractFromExcel', 'error');
					errorCount++;
				}
			}
			
			const resultMsg = `Batch extraction completed: ${successCount} successful, ${errorCount} failed`;
			log(resultMsg, 'extractFromExcel', 'success');
			vscode.window.showInformationMessage(resultMsg);
			return;
		}
		
		// Validate URI parameter - don't show file dialog for invalid input
		if (uri && (!uri.fsPath || typeof uri.fsPath !== 'string')) {
			const errorMsg = 'Invalid URI parameter provided to extractFromExcel command';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'extractFromExcel', 'error');
			return;
		}
		
		// NEVER show file dialogs - extension works only through VS Code UI
		if (!uri?.fsPath) {
			const errorMsg = 'No Excel file specified. Use right-click on an Excel file or Command Palette with file open.';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'extractFromExcel', 'error');
			return;
		}
		
		const excelFile = uri.fsPath;
		if (!excelFile) {
			log('No Excel file selected for extraction', 'extractFromExcel', 'warn');
			return;
		}

		log(`Starting Power Query extraction from: ${path.basename(excelFile)}`, 'extractFromExcel', 'info');
		vscode.window.showInformationMessage(`Extracting Power Query from: ${path.basename(excelFile)}`);
		
		// Try to use excel-datamashup for extraction
		try {
			log('Loading required modules...', 'extractFromExcel', 'debug');
			// First, we need to extract the DataMashup XML from the Excel file (scanning all customXml files)
			const JSZip = (await import('jszip')).default;
			
			// Use require for excel-datamashup to avoid ES module issues
			const excelDataMashup = require('excel-datamashup');
			log('Modules loaded successfully', 'extractFromExcel', 'debug');
			log('Reading Excel file buffer...', 'extractFromExcel', 'debug');
		let buffer: Buffer;
		try {
			buffer = fs.readFileSync(excelFile);
			const fileSizeMB = (buffer.length / (1024 * 1024)).toFixed(2);
			log(`Excel file read: ${fileSizeMB} MB`, 'extractFromExcel', 'info');
		} catch (error) {
			const errorMsg = `Failed to read Excel file: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'extractFromExcel', 'error');
			return;
		}

		log('Loading ZIP structure...', 'extractFromExcel', 'debug');
		let zip: any;
		try {
			zip = await JSZip.loadAsync(buffer, {
				checkCRC32: false // Skip CRC check for better performance on large files
			});
			log('ZIP structure loaded successfully', 'extractFromExcel', 'debug');
		} catch (error) {
			const errorMsg = `Failed to load Excel file as ZIP: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'extractFromExcel', 'error');
			return;
		}
			
			// Debug: List all files in the Excel zip
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			log(`Files in Excel archive: ${allFiles.length} total files`, 'extractFromExcel', 'info');

			// Look for Power Query DataMashup using unified detection function
			const dataMashupResults = await scanForDataMashup(zip, allFiles, undefined, false);
			const dataMashupFiles = dataMashupResults.filter(r => r.hasDataMashup);
			
			// Check for CRITICAL ISSUE: Files with <DataMashup tags but malformed structure
			const malformedDataMashupFiles = dataMashupResults.filter(r => 
				!r.hasDataMashup && 
				r.error && 
				r.error.includes('MALFORMED:')
			);
			
			if (malformedDataMashupFiles.length > 0) {
				// HARD ERROR: Found DataMashup tags but they're malformed
				const malformedFile = malformedDataMashupFiles[0];
				const errorMsg = `❌ CRITICAL ERROR: Found malformed DataMashup in ${malformedFile.file}\n\n` +
					`The file contains <DataMashup> tags but they are missing required xmlns namespace.\n` +
					`This indicates corrupted or invalid Power Query data that cannot be extracted.\n\n` +
					`Expected format: <DataMashup [sqmid="{optional-guid}"] xmlns="http://schemas.microsoft.com/DataMashup">\n` +
					`Found format: Likely missing xmlns namespace or malformed structure\n\n` +
					`Please check the Excel file's Power Query configuration.`;
				
				vscode.window.showErrorMessage(errorMsg);
				log(errorMsg, 'extractFromExcel', 'error');
				return; // HARD STOP - don't create placeholder files for malformed DataMashup
			}
			
			if (dataMashupFiles.length === 0) {
				// No DataMashup found - no actual Power Query in this file
				const customXmlFiles = allFiles.filter(f => f.startsWith('customXml/'));
				const xlFiles = allFiles.filter(f => f.startsWith('xl/') && f.includes('quer'));
				
				vscode.window.showWarningMessage(
					`No Power Query found. This Excel file does not contain DataMashup Power Query M code.\n` +
					`Available files:\n` +
					`CustomXml: ${customXmlFiles.join(', ') || 'none'}\n` +
					`Query files: ${xlFiles.join(', ') || 'none'} (these contain only metadata, not M code)\n` +
					`Total files: ${allFiles.length}`
				);
				return;
			}
			
			// Use the first DataMashup found
			const primaryDataMashup = dataMashupFiles[0];
			const foundLocation = primaryDataMashup.file;
			
			// Re-read the content for parsing (we need the actual content)
			const xmlFile = zip.file(foundLocation);
			if (!xmlFile) {
				throw new Error(`Could not re-read DataMashup file: ${foundLocation}`);
			}
			
			// Read with proper encoding detection (same logic as unified function)
			const binaryData = await xmlFile.async('nodebuffer');
			let xmlContent: string;
			
			if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
				log(`Detected UTF-16 LE BOM in ${foundLocation}`, 'extractFromExcel', 'debug');
				xmlContent = binaryData.subarray(2).toString('utf16le');
			} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
				log(`Detected UTF-8 BOM in ${foundLocation}`, 'extractFromExcel', 'debug');
				xmlContent = binaryData.subarray(3).toString('utf8');
			} else {
				xmlContent = binaryData.toString('utf8');
			}
			
			log(`Attempting to parse DataMashup Power Query from: ${foundLocation}`, 'extractFromExcel', 'debug');
			log(`DataMashup XML content size: ${(xmlContent.length / 1024).toFixed(2)} KB`, 'extractFromExcel', 'debug');
			
			// Use excel-datamashup for DataMashup format
			log('Calling excelDataMashup.ParseXml()...', 'extractFromExcel', 'debug');
			const parseResult = await excelDataMashup.ParseXml(xmlContent);
			log(`ParseXml() completed. Result type: ${typeof parseResult}`, 'extractFromExcel', 'debug');
			
			if (typeof parseResult === 'string') {
				const errorMsg = `Power Query parsing failed: ${parseResult}\nLocation: ${foundLocation}\nXML preview: ${xmlContent.substring(0, 200)}...`;
				log(errorMsg, 'extractFromExcel', 'error');
				vscode.window.showErrorMessage(errorMsg);
				return;
			}
			
			log('ParseXml() succeeded. Extracting formula...', 'extractFromExcel', 'debug');
			let formula: string;
			try {
				// Extract the formula using robust API detection
				if (typeof parseResult.getFormula === 'function') {
					formula = parseResult.getFormula();
				} else {
					// Try the module-level function
					if (typeof excelDataMashup.getFormula === 'function') {
						formula = excelDataMashup.getFormula(parseResult);
					} else {
						// Check if parseResult directly contains the formula
						formula = parseResult.formula || parseResult.code || parseResult.m;
					}
				}
				log(`getFormula() completed. Formula length: ${formula ? formula.length : 'null'}`, 'extractPowerQuery', 'debug');
			} catch (formulaError) {
				const errorMsg = `Formula extraction failed: ${formulaError}`;
				log(errorMsg, 'extractFromExcel', 'error');
				vscode.window.showErrorMessage(errorMsg);
				return;
			}
			
			if (!formula) {
				const warningMsg = `No Power Query formula found in ${foundLocation}. ParseResult keys: ${Object.keys(parseResult).join(', ')}`;
				log(warningMsg, 'extractFromExcel', 'warn');
				vscode.window.showWarningMessage(warningMsg);
				return;
			}
			
			log('Formula extracted successfully. Creating output file...', 'extractPowerQuery', 'debug');
			// Create output file with the actual formula
			const baseName = path.basename(excelFile);
			const outputPath = path.join(path.dirname(excelFile), `${baseName}_PowerQuery.m`);
			
			// Simple informational header (removed during sync)
			const informationalHeader = `// Power Query from: ${path.basename(excelFile)}
// Pathname: ${excelFile}
// Extracted: ${new Date().toISOString()}

`;

			const content = informationalHeader + formula;

			fs.writeFileSync(outputPath, content, 'utf8');
			
			// Open the created file
			const document = await vscode.workspace.openTextDocument(outputPath);
			await vscode.window.showTextDocument(document);
			
			vscode.window.showInformationMessage(`Power Query extracted to: ${path.basename(outputPath)}`);
			log(`Successfully extracted Power Query from ${path.basename(excelFile)} to ${path.basename(outputPath)}`, 'extractFromExcel', 'success');
		
		// Track this file as recently extracted to prevent immediate auto-sync
		recentExtractions.add(outputPath);
		setTimeout(() => {
			recentExtractions.delete(outputPath);
			log(`Cleared recent extraction flag for ${path.basename(outputPath)}`, 'extractFromExcel', 'debug');
		}, 2000); // Prevent auto-sync for 2 seconds after extraction
		
		// Auto-watch if enabled
		const config = getConfig();
		if (config.get<boolean>('watch.always', false)) {
			await watchFile(vscode.Uri.file(outputPath));
			log(`Auto-watch enabled for ${path.basename(outputPath)}`, 'extractPowerQuery', 'debug');
		}

		} catch (moduleError) {
			// Fallback: create a placeholder file
			const errorMsg = `Excel DataMashup parsing failed: ${moduleError}`;
			log(errorMsg, 'extractFromExcel', 'error');
			vscode.window.showWarningMessage(`${errorMsg}. Creating placeholder file for testing.`);
			
			const baseName = path.basename(excelFile); // Keep full filename including extension
			const outputPath = path.join(path.dirname(excelFile), `${baseName}_PowerQuery.m`);
			
			const placeholderContent = `// Power Query from: ${path.basename(excelFile)}
// Pathname: ${excelFile}
// Extracted: ${new Date().toISOString()}

// This is a placeholder file - actual extraction failed.
// Error: ${moduleError}
//
// File: ${excelFile}
// 
// Naming convention: Full filename + _PowerQuery.m
// Examples: 
//   MyWorkbook.xlsx -> MyWorkbook.xlsx_PowerQuery.m
//   MyWorkbook.xlsb -> MyWorkbook.xlsb_PowerQuery.m
//   MyWorkbook.xlsm -> MyWorkbook.xlsm_PowerQuery.m

let
	// Sample Power Query code structure
	Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
	#"Changed Type" = Table.TransformColumnTypes(Source,{"Column1", type text}),
	#"Filtered Rows" = Table.SelectRows(#"Changed Type", each [Column1] <> null),
	Result = #"Filtered Rows"
in
	Result`;

			fs.writeFileSync(outputPath, placeholderContent, 'utf8');
			
			// Open the created file
			const document = await vscode.workspace.openTextDocument(outputPath);
			await vscode.window.showTextDocument(document);
					vscode.window.showInformationMessage(`Placeholder file created: ${path.basename(outputPath)}`);
		log(`Created placeholder file: ${path.basename(outputPath)}`, 'extractPowerQuery', 'info');
		
		// Track this file as recently extracted to prevent immediate auto-sync
		recentExtractions.add(outputPath);
		setTimeout(() => {
			recentExtractions.delete(outputPath);
			log(`Cleared recent extraction flag for placeholder ${path.basename(outputPath)}`, 'extractFromExcel', 'debug');
		}, 2000); // Prevent auto-sync for 2 seconds after extraction
		
		// Auto-watch if enabled
		const config = getConfig();
		if (config.get<boolean>('watch.always', false)) {
			await watchFile(vscode.Uri.file(outputPath));
			log(`Auto-watch enabled for placeholder ${path.basename(outputPath)}`, 'extractPowerQuery', 'debug');
		}
		}
		
	} catch (error) {
		const errorMsg = `Failed to extract Power Query: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, 'extractFromExcel', 'error');
	}
}

/**
 * What a sync actually did.
 *
 * `syncToExcel` used to communicate by throwing or not throwing, which cannot express "some of it
 * worked". A partial live sync warned the user and returned normally, and `syncAndDelete` read that
 * as success and deleted the .m file - destroying the source while the workbook was still missing
 * queries. A destructive caller must be able to require `success` explicitly.
 */
export type SyncOutcome =
	| { status: 'success' }
	| { status: 'partial'; failures: { name: string; message: string }[] }
	| { status: 'aborted' };

async function syncToExcel(uri?: vscode.Uri, uris?: vscode.Uri[]): Promise<SyncOutcome> {
	let backupPath: string | null = null;
	
	try {
		// Handle multiple file selection (batch operations)
		if (uris && uris.length > 1) {
			log(`Batch sync started: ${uris.length} .m files selected`, 'syncToExcel', 'info');
			vscode.window.showInformationMessage(`Syncing ${uris.length} .m files to Excel...`);
			
			let successCount = 0;
			let errorCount = 0;
			
			for (const fileUri of uris) {
				try {
					await syncToExcel(fileUri); // Recursive call for single file
					successCount++;
				} catch (error) {
					log(`❌ Failed to sync ${path.basename(fileUri.fsPath)}: ${error}`, 'syncToExcel', 'error');
					errorCount++;
				}
			}
			
			const resultMsg = `Batch sync completed: ${successCount} successful, ${errorCount} failed`;
			log(resultMsg, 'syncToExcel', errorCount > 0 ? 'warn' : 'success');
			if (errorCount > 0) {
				vscode.window.showWarningMessage(`⚠️ ${resultMsg}`);
				return { status: 'partial', failures: [] };
			}
			vscode.window.showInformationMessage(`✅ ${resultMsg}`);
			return { status: 'success' };
		}
		
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			const receivedUri = uri ? `URI: ${uri.toString()}` : 'no URI provided';
			const activeFile = vscode.window.activeTextEditor?.document.fileName || 'no active file';
			throw new Error(`syncToExcel requires .m file URI. Received: ${receivedUri}, Active file: ${activeFile}`);
		}

		// Find corresponding Excel file from filename
		let excelFile = await findExcelFile(mFile);
		
		if (!excelFile) {
			// In test environment, use a test fixture or skip
			if (isTestEnvironment()) {
				const testFixtures = ['simple.xlsx', 'complex.xlsm', 'binary.xlsb'];
				for (const fixture of testFixtures) {
					const fixturePath = getTestFixturePath(fixture);
					if (fs.existsSync(fixturePath)) {
						excelFile = fixturePath;
						log(`Test environment: Using fixture ${fixture} for sync`, 'syncToExcel', 'debug');
						break;
					}
				}
				if (!excelFile) {
					log('Test environment: No Excel fixtures found, skipping sync', 'syncToExcel', 'info');
					return { status: 'aborted' };
				}
			} else {
				// SAFETY: Hard fail instead of dangerous file picker
				const mFileName = path.basename(mFile);
				const expectedExcelFile = mFileName.replace(/_PowerQuery\.m$/, '');
				
				vscode.window.showErrorMessage(
					`❌ SAFETY STOP: Cannot find corresponding Excel file.\n\n` +
					`Expected: ${expectedExcelFile}\n` +
					`Location: Same directory as ${mFileName}\n\n` +
					`To prevent accidental data destruction, please:\n` +
					`1. Ensure the Excel file is in the same directory\n` +
					`2. Verify correct naming: filename.xlsx → filename.xlsx_PowerQuery.m\n` +
					`3. Do not rename files after extraction\n\n` +
					`Extension will NOT offer to select a different file to protect your data.`
				);
				log(`SAFETY STOP: Refusing to sync ${mFileName} - corresponding Excel file not found`, 'syncToExcel', 'error');
				return { status: 'aborted' }; // HARD STOP - no file picker
			}
		}

		// Read once, up here: the lock check below now consults a setting, and so does the live
		// sync fork further down.
		const config = getConfig();

		// Check if Excel file is writable (not locked by Excel or another process)
		//
		// "Locked" is exactly the case live sync exists to handle, so ask whether we can go through
		// Excel BEFORE telling the user to close it. Without this the check fires first and returns,
		// and live sync never gets a look in - which is precisely what happened on the first real
		// attempt: the setting was on, the plumbing worked, and the user still got "close Excel".
		const isWritable = await isExcelFileWritable(excelFile);
		let liveSyncPossible = false;

		if (!isWritable && config.get<boolean>('sync.liveWhenOpen', false) && isLiveSyncSupported()) {
			const extensionPath = vscode.extensions
				.getExtension('ewc3labs.excel-power-query-editor')?.extensionPath;
			if (extensionPath) {
				const probe = await getLiveStatus(excelFile, extensionPath);
				liveSyncPossible = probe.open;

				// REFUSE TO WRITE OVER UNSAVED WORK.
				//
				// The backup taken before a sync is a copy of the file on DISK - the last saved state. If
				// Excel is holding edits the user has not saved, those edits exist in no file anywhere, and
				// a live write replaces them with no way back. The backup is worse than absent here: it
				// looks like protection and is not.
				//
				// Saving on their behalf is not ours to do - it would commit changes they may still have
				// been deciding about. Say what is in the way and let them choose.
				if (config.get<boolean>('sync.requireSavedWorkbook', true)
					&& shouldRefuseUnsavedWorkbook(probe)) {
					const message = `${path.basename(excelFile)} has unsaved changes in Excel. Save it there `
						+ 'first - a backup can only capture what is on disk, so syncing now would write over '
						+ 'work that nothing has a copy of.';
					log(message, 'syncToExcel', 'warn');
					vscode.window.showWarningMessage(message);
					return { status: 'aborted' };
				}
				if (probe.open && probe.saved === false) {
					log('Workbook has unsaved changes and sync.requireSavedWorkbook is off - '
						+ 'writing anyway. The backup holds the last SAVED state, not what is in Excel.',
						'syncToExcel', 'warn');
				}

				log(`Excel file is locked; live sync ${probe.open ? 'CAN' : 'cannot'} handle it ` +
					`(available=${probe.available}${probe.reason ? ', ' + probe.reason : ''}` +
					`${probe.excelProcesses ? ', excelProcesses=' + probe.excelProcesses : ''})`,
					'syncToExcel', 'info');

				// The file is locked AND Excel is running AND we cannot see the workbook. That is
				// not "not open" - something is hiding it, and the user deserves to know what.
				const explanation = explainInvisibleWorkbook(probe);
				if (explanation) {
					log(explanation, 'syncToExcel', 'warn');
					vscode.window.showWarningMessage(explanation);
					return { status: 'aborted' };
				}
			}
		}

		if (!isWritable && !liveSyncPossible) {
			const fileName = path.basename(excelFile);

			// We KNOW it is locked - the write access probe failed. What we cannot know is by whom,
			// because COM partitions running objects by integrity level: a workbook open in an
			// elevated Excel is invisible to a normal-integrity extension, and the reverse. Say that
			// plainly rather than "possibly open in Excel", and name both ways out. Guessing at a
			// similarly-named workbook we CAN see is how queries end up in the wrong file.
			const message = explainLockedButUnreachable(excelFile);
			log(`${fileName} is locked and not reachable through Excel `
				+ `(owner lock file present: ${hasOwnerLockFile(excelFile)})`, 'syncToExcel', 'warn');

			const retry = await vscode.window.showWarningMessage(message, 'Retry', 'Cancel');
			if (retry === 'Retry') {
				// Retry after a short delay
				setTimeout(() => syncToExcel(uri), 1000);
			}
			return { status: 'aborted' };
		}

		// Read the .m file content
		const mContent = fs.readFileSync(mFile, 'utf8');
		
	// Extract just the M code - find the section declaration and discard everything above it
	// DataMashup content always starts with "section <SectionName>;"
	const sectionMatch = mContent.match(/^(.*?)(section\s+\w+\s*;[\s\S]*)$/m);
	
	let cleanMCode;
	if (sectionMatch) {
		// Found section declaration - use everything from section onwards
		cleanMCode = sectionMatch[2].trim();
		const headerLength = sectionMatch[1].length;
		log(`Header stripping - Found section at position ${headerLength}, removed ${headerLength} header characters`, 'syncToExcel', 'verbose');
	} else {
		// No section found - use original content (might be a different format)
		cleanMCode = mContent.trim();
		log(`Header stripping - No section declaration found, using original content`, 'syncToExcel', 'debug');
	}
		
		if (!cleanMCode) {
			vscode.window.showErrorMessage('No Power Query M code found in file.');
			return { status: 'aborted' };
		}
		
		// Create backup of Excel file if enabled
		
		if (config.get<boolean>('backup.autoBackupBeforeSync', true)) {
			const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
			backupPath = getBackupPath(excelFile, timestamp);
			
			// Ensure backup directory exists
			const backupDir = path.dirname(backupPath);
			if (!fs.existsSync(backupDir)) {
				fs.mkdirSync(backupDir, { recursive: true });
			}
			
			fs.copyFileSync(excelFile, backupPath);
			vscode.window.showInformationMessage(`Syncing to Excel... (Backup created: ${path.basename(backupPath)})`);
			log(`Backup created: ${backupPath}`, 'syncToExcel', 'verbose');
			
			// Clean up old backups
			cleanupOldBackups(excelFile);
		} else {
			vscode.window.showInformationMessage(`Syncing to Excel... (No backup - disabled in settings)`);
		}
		
		// --- live sync -----------------------------------------------------------------------
		// If Excel already has this workbook open, the zip below cannot be written: Excel holds an
		// exclusive WRITE lock. Historically that meant failing, or waiting for the user to close
		// the file. Instead, ask the running Excel to make the change through its own object model.
		//
		// The backup above still ran, deliberately. A live write leaves the workbook dirty rather
		// than changing the file, but the user may well save it - and then the on-disk state has
		// changed with no backup, which is not a trade this project makes.
		if (config.get<boolean>('sync.liveWhenOpen', false) && isLiveSyncSupported()) {
			const extensionPath = vscode.extensions
				.getExtension('ewc3labs.excel-power-query-editor')?.extensionPath;

			if (!extensionPath) {
				log('Live sync enabled but the extension path is unavailable; using the on-disk writer',
					'syncToExcel', 'warn');
			} else {
				const status = await getLiveStatus(excelFile, extensionPath);

				if (!status.available) {
					// Not a failure. Excel is not running, or we could not ask - either way the file
					// is not locked by it, so the normal path below is the right answer.
					log(`Live sync unavailable (${status.reason}); using the on-disk writer`,
						'syncToExcel', 'debug');
				} else if (!status.open) {
					log('Workbook is not open in Excel; using the on-disk writer', 'syncToExcel', 'debug');
				} else {
					const section = parseSection(cleanMCode);
					if (section.queries.length === 0) {
						log('No shared bindings found in the .m file; using the on-disk writer',
							'syncToExcel', 'warn');
					} else {
						const diff = diffQueries(section, status.queries);
						const payload = [...diff.update, ...diff.add]
							.map(q => ({ name: q.name, formula: q.expression }));

						log(`Live sync: ${diff.update.length} to update, ${diff.add.length} to add, ` +
							`${diff.missingFromDocument.length} in the workbook but not the file`,
							'syncToExcel', 'info');

						const result = await writeLive(excelFile, payload, extensionPath);

						if (!result.ok && result.reason) {
							throw new Error(`Live sync failed: ${result.reason}`);
						}
						if (result.failures.length > 0) {
							const detail = result.failures.map(f => `${f.name}: ${f.message}`).join('; ');
							vscode.window.showWarningMessage(
								`Synced to the open workbook, but ${result.failures.length} query(s) FAILED: ${detail}`);
						}

						const parts: string[] = [];
						if (result.updated.length) { parts.push(`${result.updated.length} updated`); }
						if (result.added.length) { parts.push(`${result.added.length} added`); }
						if (result.unchanged.length) { parts.push(`${result.unchanged.length} unchanged`); }

						// Say plainly that the file on disk has NOT changed. A user who has been
						// trained by every previous version to expect a written file needs to know
						// their workbook now has unsaved changes.
						// Tell the truth per workbook, because the two cases genuinely differ. MEASURED:
						// with AutoSave on, a write through COM is committed to disk within about two
						// seconds and closing without saving does NOT undo it. Promising a review step
						// there would be false, and false in the direction that loses work.
						const saveNote = !result.dirty && !result.autoSaveOn
							? ' — nothing needed changing.'
							: result.autoSaveOn
								? ' — AutoSave is on, so Excel saves this itself within seconds. To undo, '
								  + 'use version history in OneDrive or SharePoint.'
								: ' — the workbook now has unsaved changes in Excel, save it there.';

						// A sync where some queries failed is NOT a success, and must not be announced
						// as one. The warning above names the failures, but the message that arrives
						// second is the one that stays on screen and the level that gets logged is the
						// one someone triaging an issue will read. Both have to tell the same story.
						const partial = result.failures.length > 0;
						const summary = partial
							? `Synced to the open workbook with ${result.failures.length} failure(s) ` +
							  `(${parts.join(', ') || 'no changes'})${saveNote}`
							: `Synced to the open workbook (${parts.join(', ') || 'no changes'})${saveNote}`;

						log(summary, 'syncToExcel', partial ? 'warn' : 'success');
						if (!partial) {
							vscode.window.showInformationMessage(summary);
						}

						if (diff.missingFromDocument.length > 0) {
							// Never deleted, only reported - see diffQueries.
							const names = diff.missingFromDocument.join(', ');
							log(`Left alone (present in the workbook, absent from the .m file): ${names}`,
								'syncToExcel', 'info');
						}
						return partial
							? { status: 'partial', failures: result.failures }
							: { status: 'success' };
					}
				}
			}
		}

		// Load Excel file as ZIP
		const JSZip = (await import('jszip')).default;
		const xml2js = await import('xml2js');
		const excelDataMashup = require('excel-datamashup');
		
		const buffer = fs.readFileSync(excelFile);
		const zip = await JSZip.loadAsync(buffer);
		
		// Find the DataMashup XML file by scanning all customXml files
		const customXmlFiles = Object.keys(zip.files)
			.filter(name => name.startsWith('customXml/') && name.endsWith('.xml'))
			.filter(name => !name.includes('/_rels/')) // Exclude relationship files
			.sort();
		
		// Find the DataMashup XML file 
		// NOTE: Metadata parsing not implemented - scan all customXml files
		let dataMashupFile = null;
		let dataMashupLocation = '';
		
		// Scan customXml files for DataMashup content using efficient detection
		for (const location of customXmlFiles) {
			const file = zip.file(location);
			if (file) {
				try {
					// Use same binary reading and BOM handling as extraction
					const binaryData = await file.async('nodebuffer');
					let content: string;
					
					// Check for UTF-16 LE BOM (FF FE)
					if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
						content = binaryData.subarray(2).toString('utf16le');
					} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
						content = binaryData.subarray(3).toString('utf8');
					} else {
						content = binaryData.toString('utf8');
					}
					
					// Quick pre-filter: only check files that contain DataMashup opening tag
					if (!content.includes('<DataMashup')) {
						continue; // Skip silently
					}
					
					// IMPROVED DataMashup detection - look for actual DataMashup XML structure
					const hasDataMashupOpenTag = /<DataMashup(\s+sqmid=".+?")?\s+xmlns="http:\/\/schemas\.microsoft\.com\/DataMashup">/.test(content);
					const hasDataMashupCloseTag = content.includes('</DataMashup>');
					const isSchemaRefOnly = content.includes('ds:schemaRef') && content.includes('http://schemas.microsoft.com/DataMashup');
					
					if (hasDataMashupOpenTag && hasDataMashupCloseTag && !isSchemaRefOnly) {
						dataMashupFile = file;
						dataMashupLocation = location;
						log(`Found DataMashup content for sync in: ${location}`, 'syncToExcel', 'debug');
						break; // Found it!
					}
					// All other cases: skip silently (no logging for schema refs or malformed content)
				} catch (e) {
					log(`Could not check ${location}: ${e}`, 'syncToExcel', 'warn');
				}
			}
		}
		
		if (!dataMashupFile) {
			vscode.window.showErrorMessage('No DataMashup found in Excel file. This file may not contain Power Query.');
			return { status: 'aborted' };
		}
		
		// Read and decode the DataMashup XML
		const binaryData = await dataMashupFile.async('nodebuffer');
		let dataMashupXml: string;
		
		// Handle UTF-16 LE BOM like in extraction
		if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
			log('Detected UTF-16 LE BOM in DataMashup', 'syncToExcel', 'debug');
			dataMashupXml = binaryData.subarray(2).toString('utf16le');
		} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
			log('Detected UTF-8 BOM in DataMashup', 'syncToExcel', 'debug');
			dataMashupXml = binaryData.subarray(3).toString('utf8');
		} else {
			dataMashupXml = binaryData.toString('utf8');
		}
		
		if (!dataMashupXml.includes('DataMashup')) {
			vscode.window.showErrorMessage('Invalid DataMashup format in Excel file.');
			return { status: 'aborted' };
		}
		
		// Debug code removed: No longer saving original DataMashup XML when logLevel is 'debug'.
		// // DEBUG: Save the original DataMashup XML for inspection (debug mode only)
		// const logLevel = getConfig().get<string>('log.level', 'info');
		// if (logLevel === 'debug') {
		// 	const baseName = path.basename(excelFile, path.extname(excelFile));
		// 	const debugDir = path.join(path.dirname(excelFile), `${baseName}_sync_debug`);
		// 	if (!fs.existsSync(debugDir)) {
		// 		fs.mkdirSync(debugDir, { recursive: true });
		// 	}
		// 	fs.writeFileSync(
		// 		path.join(debugDir, 'original_datamashup.xml'),
		// 		dataMashupXml,
		// 		'utf8'
		// 	);
		// 	log(`Debug: Saved original DataMashup XML to ${path.basename(debugDir)}/original_datamashup.xml`, 'syncToExcel', 'debug');
		// }
		// Debug code removed: No longer saving original DataMashup XML when logLevel is 'debug'.

		
		// Use excel-datamashup to correctly update the DataMashup binary content
		try {
			log('Attempting to parse existing DataMashup with excel-datamashup...', 'syncToExcel', 'debug');
			// Parse the existing DataMashup to get structure
			const parseResult = await excelDataMashup.ParseXml(dataMashupXml);
			
			if (typeof parseResult === 'string') {
				throw new Error(`Failed to parse existing DataMashup: ${parseResult}`);
			}
			
			log('DataMashup parsed successfully, updating formula...', 'syncToExcel', 'debug');
			// Use setFormula to update the M code (this also calls resetPermissions)
			parseResult.setFormula(cleanMCode);
			
			log('Formula updated, generating new DataMashup content...', 'syncToExcel', 'debug');
			// Use save to get the updated base64 binary content
			const newBase64Content = await parseResult.save();
			
			log(`excel-datamashup save() returned type: ${typeof newBase64Content}, length: ${String(newBase64Content).length}`, 'syncToExcel', 'debug');
			
			if (typeof newBase64Content === 'string' && newBase64Content.length > 0) {
				log('✅ excel-datamashup approach succeeded, updating Excel file...', 'syncToExcel', 'debug');
				// Success! Now we need to reconstruct the full DataMashup XML with new base64 content
				// Replace the base64 content inside the DataMashup tags
				const dataMashupRegex = /<DataMashup[^>]*>(.*?)<\/DataMashup>/s;
				const newDataMashupXml = dataMashupXml.replace(dataMashupRegex, (match, oldContent) => {
					// Keep the DataMashup tag attributes but replace the base64 content
					const tagMatch = match.match(/<DataMashup[^>]*>/);
					const openingTag = tagMatch ? tagMatch[0] : '<DataMashup>';
					return `${openingTag}${newBase64Content}</DataMashup>`;
				});
				
				// Convert back to UTF-16 LE with BOM if original was UTF-16
				let newBinaryData: Buffer;
				if (binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
					// Add UTF-16 LE BOM and encode
					const utf16Buffer = Buffer.from(newDataMashupXml, 'utf16le');
					const bomBuffer = Buffer.from([0xFF, 0xFE]);
					newBinaryData = Buffer.concat([bomBuffer, utf16Buffer]);
				} else {
					// Keep as UTF-8
					newBinaryData = Buffer.from(newDataMashupXml, 'utf8');
				}
				
				// Update the ZIP with new DataMashup at the correct location
				zip.file(dataMashupLocation, newBinaryData);
				
				// Write the updated Excel file
				const updatedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
				fs.writeFileSync(excelFile, updatedBuffer);
				
				vscode.window.showInformationMessage(`✅ Successfully synced Power Query to Excel: ${path.basename(excelFile)}`);
				log(`Successfully synced Power Query to Excel: ${path.basename(excelFile)}`, 'syncToExcel', 'success');
				
				// Open Excel after sync if enabled
				const config = getConfig();
				if (config.get<boolean>('sync.openExcelAfterWrite', false)) {
					try {
						await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(excelFile));
						log(`Opened Excel file after sync: ${path.basename(excelFile)}`, 'syncToExcel', 'verbose');
					} catch (openError) {
						log(`Failed to open Excel file after sync: ${openError}`, 'syncToExcel', 'error');
					}
				}
				// The file on disk was rewritten and verified. This is the closed-workbook success path.
				return { status: 'success' };
				
			} else {
				throw new Error(`excel-datamashup save() returned invalid content - Type: ${typeof newBase64Content}, Length: ${String(newBase64Content).length}`);
			}
			
		} catch (dataMashupError) {
			log(`excel-datamashup approach failed: ${dataMashupError}`, 'syncToExcel', 'error');
			throw new Error(`DataMashup sync failed: ${dataMashupError}. The DataMashup format may have changed or be unsupported.`);
		}
		
	} catch (error) {
		const errorMsg = `Failed to sync to Excel: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(`Sync error: ${error}`, 'syncToExcel', 'error');
		
		// If we have a backup, offer to restore it
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (mFile && backupPath && fs.existsSync(backupPath)) {
			const restore = await vscode.window.showErrorMessage(
				'Sync failed. Restore from backup?',
				'Restore', 'Keep Current'
			);
			if (restore === 'Restore') {
				const excelFile = await findExcelFile(mFile);
				if (excelFile) {
					fs.copyFileSync(backupPath, excelFile);
					vscode.window.showInformationMessage('Excel file restored from backup.');
					log(`Restored from backup: ${backupPath}`, 'syncToExcel', 'info');
				}
			}
		}
		// The error was reported and, where possible, the workbook restored. Nothing synced.
		return { status: 'aborted' };
	}
}

async function watchFile(uri?: vscode.Uri, uris?: vscode.Uri[]): Promise<void> {
	try {
		// Handle multiple file selection (batch operations)
		if (uris && uris.length > 1) {
			log(`Batch watch started: ${uris.length} .m files selected`, 'watchFile', 'info');
			vscode.window.showInformationMessage(`Setting up watchers for ${uris.length} .m files...`);
			
			let successCount = 0;
			let errorCount = 0;
			
			for (const fileUri of uris) {
				try {
					await watchFile(fileUri); // Recursive call for single file
					successCount++;
				} catch (error) {
					log(`Failed to watch ${path.basename(fileUri.fsPath)}: ${error}`, 'watchFile', 'error');
					errorCount++;
				}
			}
			
			const resultMsg = `Batch watch completed: ${successCount} successful, ${errorCount} failed`;
			log(resultMsg, 'watchFile', 'success');
			vscode.window.showInformationMessage(resultMsg);
			return;
		}
		
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			const receivedUri = uri ? `URI: ${uri.toString()}` : 'no URI provided';
			const activeFile = vscode.window.activeTextEditor?.document.fileName || 'no active file';
			throw new Error(`watchFile requires .m file URI. Received: ${receivedUri}, Active file: ${activeFile}`);
		}

		if (fileWatchers.has(mFile)) {
			vscode.window.showInformationMessage(`File is already being watched: ${path.basename(mFile)}`);
			return;
		}

		// Verify that corresponding Excel file exists
		const excelFile = await findExcelFile(mFile);
		if (!excelFile) {
			// In test environment, proceed without user interaction
			if (isTestEnvironment()) {
				log('Test environment: Missing Excel file, proceeding with watch anyway', 'watchFile', 'info');
			} else {
				const selection = await vscode.window.showWarningMessage(
					`Cannot find corresponding Excel file for ${path.basename(mFile)}. Watch anyway?`,
					'Yes, Watch Anyway', 'No'
				);
				if (selection !== 'Yes, Watch Anyway') {
					return;
				}
			}
		}

	// Debug logging for watcher setup
	log(`Setting up file watcher for: ${mFile}`, 'watchFile', 'info');
	log(`Remote environment: ${vscode.env.remoteName}`, 'watchFile', 'verbose');
	log(`Is dev container: ${vscode.env.remoteName === 'dev-container'}`, 'watchFile', 'verbose');
	
	const isDevContainer = vscode.env.remoteName === 'dev-container';
	
	// PRIMARY WATCHER: Always use Chokidar as the main watcher
	const watcher = watch(mFile, { 
		ignoreInitial: true,
		usePolling: isDevContainer, // Use polling in dev containers for better compatibility
		interval: isDevContainer ? 1000 : undefined, // Poll every second in dev containers
		awaitWriteFinish: {
			stabilityThreshold: 300,
			pollInterval: 100
		}
	});
	
	log(`CHOKIDAR watcher created for ${path.basename(mFile)}, polling: ${isDevContainer}`, 'watchFile', 'verbose');
	
	// Add comprehensive event logging
	watcher.on('change', async () => {			try {
				log(`CHOKIDAR: File change detected: ${path.basename(mFile)}`, 'watchFile', 'verbose');
				vscode.window.showInformationMessage(`📝 File changed, syncing: ${path.basename(mFile)}`);
				log(`File changed, triggering debounced sync: ${path.basename(mFile)}`, 'watchFile', 'verbose');
				debouncedSyncToExcel(mFile).catch(error => {
					const errorMsg = `Auto-sync failed: ${error}`;
					vscode.window.showErrorMessage(errorMsg);
					log(errorMsg, 'watchFile', 'error');
				});
			} catch (error) {
				const errorMsg = `Auto-sync failed: ${error}`;
				vscode.window.showErrorMessage(errorMsg);
				log(errorMsg, 'watchFile', 'error');
			}
	});
	
	watcher.on('add', (path) => {
		log(`CHOKIDAR: File added: ${path}`, 'watchFile', 'info');
		// DON'T trigger sync on file creation - only on user changes
	});
	
	watcher.on('unlink', (path) => {
		log(`CHOKIDAR: File deleted: ${path}`, 'watchFile', 'info');
	});
	
	watcher.on('error', (error) => {
		log(`CHOKIDAR: Watcher error: ${error}`, 'watchFile', 'error');
	});
	
	watcher.on('ready', () => {
		log(`CHOKIDAR: Watcher ready for ${path.basename(mFile)}`, 'watchFile', 'info');
	});

	// BACKUP WATCHER: Only add VS Code FileSystemWatcher in dev containers as backup
	let vscodeWatcher: vscode.FileSystemWatcher | undefined;
	let documentWatcher: vscode.Disposable | undefined;
	
	if (isDevContainer) {
		log(`Adding backup watchers for dev container environment`, 'watchFile', 'verbose');
		
		vscodeWatcher = vscode.workspace.createFileSystemWatcher(mFile);
		vscodeWatcher.onDidChange(async () => {
			try {
				log(`VSCODE: File change detected: ${path.basename(mFile)}`, 'watchFile', 'info');
				vscode.window.showInformationMessage(`📝 File changed (VSCode watcher), syncing: ${path.basename(mFile)}`);
				debouncedSyncToExcel(mFile).catch(error => {
					log(`VS Code watcher sync failed: ${error}`, 'watchFile', 'info');
				});
			} catch (error) {
				log(`VS Code watcher sync failed: ${error}`, 'watchFile', 'info');
			}
		});
		
		vscodeWatcher.onDidCreate(() => {
			log(`VSCODE: File created: ${path.basename(mFile)}`, 'watchFile', 'info');
		});
		
		vscodeWatcher.onDidDelete(() => {
			log(`VSCODE: File deleted: ${path.basename(mFile)}`, 'watchFile', 'info');
		});

		log(`VS Code FileSystemWatcher created for ${path.basename(mFile)}`, 'watchFile', 'info');

		// EXPERIMENTAL: Document save events as additional trigger (dev container only)
		documentWatcher = vscode.workspace.onDidSaveTextDocument(async (document) => {
			if (document.fileName === mFile) {
				try {
					log(`documentWatcher: Save event detected: ${path.basename(mFile)}`, 'watchFile', 'verbose');
					vscode.window.showInformationMessage(`📝 File saved (document event), syncing: ${path.basename(mFile)}`);
					debouncedSyncToExcel(mFile).catch(error => {
						log(`documentWatcher: Save event sync failed: ${error}`, 'watchFile', 'error');
					});
				} catch (error) {
					log(`documentWatcher: Save event sync failed: ${error}`, 'watchFile', 'error');
				}
			}
		});
		
		log(`VS Code document save watcher created for ${path.basename(mFile)}`, 'watchFile', 'info');
	} else {
		log(`Windows environment detected - using Chokidar only to avoid cascade events`, 'watchFile', 'verbose');
	}		// Store watchers for cleanup (handle optional backup watchers)
		const watcherSet = { 
			chokidar: watcher, 
			vscode: vscodeWatcher || null,
			document: documentWatcher || null
		};
		fileWatchers.set(mFile, watcherSet);
		
		const excelFileName = excelFile ? path.basename(excelFile) : 'Excel file (when found)';
		vscode.window.showInformationMessage(`👀 Now watching: ${path.basename(mFile)} → ${excelFileName}`);
		log(`Started watching: ${path.basename(mFile)}`, 'watch', 'info');
		updateStatusBar();
		
		// Ensure the Promise resolves after watchers are set up
		return Promise.resolve();
		
	} catch (error) {
		const errorMsg = `Failed to watch file: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(`Watch error: ${error}`, 'watchFile', 'error');
	}
}

async function toggleWatch(uri?: vscode.Uri): Promise<void> {
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			const receivedUri = uri ? `URI: ${uri.toString()}` : 'no URI provided';
			const activeFile = vscode.window.activeTextEditor?.document.fileName || 'no active file';
			throw new Error(`toggleWatch requires .m file URI. Received: ${receivedUri}, Active file: ${activeFile}`);
		}

		const isWatching = fileWatchers.has(mFile);
		
		if (isWatching) {
			// Stop watching
			await stopWatching(uri);
		} else {
			// Start watching
			await watchFile(uri);
		}
		
	} catch (error) {
		const errorMsg = `Failed to toggle watch: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, 'toggleWatch', 'verbose');
		log(`Toggle watch error: ${error}`, 'toggleWatch', 'error');
	}
}

async function stopWatching(uri?: vscode.Uri): Promise<void> {
	const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
	if (!mFile) {
		return;
	}

	const watchers = fileWatchers.get(mFile);
	if (watchers) {
		await watchers.chokidar.close();
		watchers.vscode?.dispose();
		watchers.document?.dispose();
		fileWatchers.delete(mFile);
		vscode.window.showInformationMessage(`Stopped watching: ${path.basename(mFile)}`);
		log(`Stopped watching: ${path.basename(mFile)}`, 'stopWatching', 'verbose');
		updateStatusBar();
	} else {
		vscode.window.showInformationMessage(`File was not being watched: ${path.basename(mFile)}`);
	}
}

async function syncAndDelete(uri?: vscode.Uri): Promise<void> {
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			const receivedUri = uri ? `URI: ${uri.toString()}` : 'no URI provided';
			const activeFile = vscode.window.activeTextEditor?.document.fileName || 'no active file';
			throw new Error(`syncAndDelete requires .m file URI. Received: ${receivedUri}, Active file: ${activeFile}`);
		}

		const config = getConfig();
		let confirmation: string | undefined = 'Yes, Sync & Delete';
		
		// Ask for confirmation if setting is enabled
		if (config.get<boolean>('sync.deleteAlwaysConfirm', true)) {
			confirmation = await vscode.window.showWarningMessage(
				`Sync ${path.basename(mFile)} to Excel and then delete the .m file?`,
				{ modal: true },
				'Yes, Sync & Delete', 'Cancel'
			);
		}
		
		if (confirmation === 'Yes, Sync & Delete') {
			// First try to sync
			try {
				// REQUIRE SUCCESS BEFORE DELETING ANYTHING.
				//
				// syncToExcel used to say "it worked" by not throwing, and a partial live sync does not
				// throw - it warns and returns. This deleted the .m file while the workbook was still
				// missing queries, destroying the source and leaving the destination incomplete. The
				// only safe reading of a destructive command is an explicit success.
				const outcome = await syncToExcel(uri);
				if (outcome.status !== 'success') {
					const detail = outcome.status === 'partial' && outcome.failures.length
						? ` (${outcome.failures.map(f => f.name).join(', ')} failed)`
						: '';
					const message = `Not deleting ${path.basename(mFile)}: the sync did not fully succeed${detail}.`;
					vscode.window.showWarningMessage(message);
					log(message, 'syncAndDelete', 'warn');
					return;
				}

				// Stop watching if enabled and if being watched
				const watchers = fileWatchers.get(mFile);
				if (watchers) {
					// `watch.offOnDelete` already means "stop watching a .m that is deleted", and Sync &
					// Delete deletes it. This used to read `syncDeleteTurnsWatchOff`, which was never
					// declared in the manifest - so it could not be set from the settings UI and always
					// took its default. One registered setting, one meaning.
					if (config.get<boolean>('watch.offOnDelete', true)) {
						await watchers.chokidar.close();
						watchers.vscode?.dispose();
						watchers.document?.dispose();
						fileWatchers.delete(mFile);
						log(`Stopped watching due to sync & delete: ${path.basename(mFile)}`, 'syncAndDelete', 'verbose');
						updateStatusBar();
					}
				}
				
				// Close the file in VS Code if it's open
				const openEditors = vscode.window.visibleTextEditors;
				for (const editor of openEditors) {
					if (editor.document.fileName === mFile) {
						await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
						break;
					}
				}
				
				// Delete the file
				fs.unlinkSync(mFile);
				vscode.window.showInformationMessage(`✅ Synced and deleted: ${path.basename(mFile)}`);
				log(`Successfully synced and deleted: ${path.basename(mFile)}`, 'syncAndDelete', 'success');
				
			} catch (syncError) {
				const errorMsg = `Sync failed, file not deleted: ${syncError}`;
				vscode.window.showErrorMessage(errorMsg);
				log(errorMsg, 'syncAndDelete', 'error');
			}
		}
	} catch (error) {
		const errorMsg = `Sync and delete failed: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(`Sync and delete error: ${error}`, 'syncAndDelete', 'error');
	}
}

// Unified DataMashup detection function used by both main extraction and debug extraction
interface DataMashupScanResult {
	file: string;
	hasDataMashup: boolean;
	size: number;
	error?: string;
	extractedFormula?: string;
}

async function scanForDataMashup(
	zip: any, 
	allFiles: string[], 
	outputDir?: string, 
	isDebugMode: boolean = false
): Promise<DataMashupScanResult[]> {
	const results: DataMashupScanResult[] = [];
	
	// Focus on customXml files first (where DataMashup actually lives)
	const customXmlFiles = allFiles
		.filter(name => name.startsWith('customXml/') && name.endsWith('.xml'))
		.filter(name => !name.includes('/_rels/')) // Exclude relationship files
		.sort(); // Process in consistent order
	
	// Only in debug mode, also scan other XML files for comparison
	const xmlFilesToScan = isDebugMode ? 
		allFiles.filter(f => f.toLowerCase().endsWith('.xml')) : 
		customXmlFiles;
	
	log(`Scanning ${xmlFilesToScan.length} XML files for DataMashup content...`, 'scanForDataMashup', 'verbose');
	
	for (const fileName of xmlFilesToScan) {
		try {
			const file = zip.file(fileName);
			if (file) {
				// Read as binary first, then decode properly (same as main extraction)
				const binaryData = await file.async('nodebuffer');
				let content: string;
				
				// Check for UTF-16 LE BOM (FF FE)
				if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
					log(`Detected UTF-16 LE BOM in ${fileName}`, 'scanForDataMashup', 'verbose');
					// Decode UTF-16 LE (skip the 2-byte BOM)
					content = binaryData.subarray(2).toString('utf16le');
				} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
					log(`Detected UTF-8 BOM in ${fileName}`, 'scanForDataMashup', 'verbose');
					// Decode UTF-8 (skip the 3-byte BOM)
					content = binaryData.subarray(3).toString('utf8');
				} else {
					// Try UTF-8 first (most common)
					content = binaryData.toString('utf8');
				}
				
				// Quick pre-filter: only process files that actually contain DataMashup opening tag
				if (!content.includes('<DataMashup')) {
					// Skip silently - no DataMashup content at all
					results.push({
						file: fileName,
						hasDataMashup: false,
						size: content.length
					});
					continue;
				}
				
				// Found <DataMashup tag - this is a legitimate candidate, so start logging
				log(`Found <DataMashup tag in ${fileName} (${(content.length / 1024).toFixed(1)} KB) - validating structure...`, 'scanForDataMashup', 'verbose');

				// IMPROVED DataMashup detection - look for actual DataMashup XML structure
				let hasDataMashup = false;
				let parseResult: any = null;
				let parseError: string | undefined;
				
				// Smart detection: look for proper DataMashup XML structure
				// Real DataMashup: <DataMashup [sqmid="{guid}"] xmlns="http://schemas.microsoft.com/DataMashup">{encoded-content}</DataMashup>
				// Schema ref only: <ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/>
				const hasDataMashupOpenTag = /<DataMashup(\s+sqmid=".+?")?\s+xmlns="http:\/\/schemas\.microsoft\.com\/DataMashup">/.test(content);
				const hasDataMashupCloseTag = content.includes('</DataMashup>');
				const isSchemaRefOnly = content.includes('ds:schemaRef') && content.includes('http://schemas.microsoft.com/DataMashup');
				
				if (hasDataMashupOpenTag && hasDataMashupCloseTag && !isSchemaRefOnly) {
					log(`Valid DataMashup XML structure detected - attempting to parse...`, 'scanForDataMashup', 'verbose');
					// This looks like actual DataMashup content - try to parse it
					try {
						const excelDataMashup = require('excel-datamashup');
						parseResult = await excelDataMashup.ParseXml(content);
						
						if (typeof parseResult === 'object' && parseResult !== null) {
							hasDataMashup = true;
							log(`Successfully parsed DataMashup content`, 'scanForDataMashup', 'success');
						} else {
							log(`ParseXml() failed: ${parseResult}`, 'scanForDataMashup', 'error');
							parseError = `Parse failed: ${parseResult}`;
						}
					} catch (error) {
						log(`Error parsing DataMashup: ${error}`, 'scanForDataMashup', 'error');
						parseError = `Parse error: ${error}`;
					}
				} else if (isSchemaRefOnly) {
					log(`Contains only DataMashup schema reference, not actual content`, 'scanForDataMashup', 'debug');
				} else if (!hasDataMashupOpenTag) {
					log(`Contains <DataMashup but missing required xmlns namespace or malformed structure`, 'scanForDataMashup', 'debug');
					parseError = 'MALFORMED: missing xmlns namespace or malformed structure';
				} else if (!hasDataMashupCloseTag) {
					log(`Contains <DataMashup opening but missing closing </DataMashup> tag`, 'scanForDataMashup', 'debug');
					parseError = 'MALFORMED: missing closing </DataMashup> tag';
				} else {
					log(`Contains <DataMashup but structure validation failed`, 'scanForDataMashup', 'debug');
					parseError = 'MALFORMED: structure validation failed';
				}
				
				const result: DataMashupScanResult = {
					file: fileName,
					hasDataMashup,
					size: content.length,
					...(parseError && { error: parseError })
				};
				
				if (hasDataMashup && parseResult) {
					// In debug mode, extract files and the M code
					if (isDebugMode && outputDir) {
						// Extract the DataMashup content to a separate file
						const safeName = fileName.replace(/[\/\\]/g, '_');
						const dataMashupPath = path.join(outputDir, `DATAMASHUP_${safeName}`);
						fs.writeFileSync(dataMashupPath, content, 'utf8');
						log(`DataMashup extracted to: ${path.basename(dataMashupPath)}`, 'scanForDataMashup', 'success');
						
						// Extract the M code using the correct API
						try {
							// Try both possible API patterns
							let formula: string | undefined;
							
							if (typeof parseResult.getFormula === 'function') {
								formula = parseResult.getFormula();
							} else {
								const excelDataMashup = require('excel-datamashup');
								if (typeof excelDataMashup.getFormula === 'function') {
									formula = excelDataMashup.getFormula(parseResult);
								} else {
									// Check if parseResult directly contains the formula
									formula = parseResult.formula || parseResult.code || parseResult.m;
								}
							}
							
							if (formula && typeof formula === 'string') {
								result.extractedFormula = formula;
								
								// Save the extracted M code
								const baseName = path.basename(fileName, '.xml');
								const mCodePath = path.join(outputDir, `${baseName}_PowerQuery.m`);
								const header = `// Power Query from: ${fileName}\n// Extracted: ${new Date().toISOString()}\n\n`;
								fs.writeFileSync(mCodePath, header + formula, 'utf8');
								log(`Extracted M code to: ${path.basename(mCodePath)} (${(formula.length / 1024).toFixed(1)} KB)`, 'scanForDataMashup', 'success');
							} else {
								log(`Could not extract formula from parseResult for ${fileName}`, 'scanForDataMashup', 'error');
								result.error = (result.error || '') + ' | Formula extraction failed';
							}
						} catch (formulaError) {
							log(`Error extracting formula from ${fileName}: ${formulaError}`, 'scanForDataMashup', 'error');
							result.error = (result.error || '') + ` | Formula error: ${formulaError}`;
						}
					}
				}
				
				results.push(result);
				
				// In debug mode, extract customXml files regardless for inspection
				if (isDebugMode && outputDir && fileName.startsWith('customXml/')) {
					const safeName = fileName.replace(/[\/\\]/g, '_');
					fs.writeFileSync(
						path.join(outputDir, `${safeName}.txt`),
						content,
						'utf8'
					);
				}
			}
		} catch (error) {
			log(`Error scanning ${fileName}: ${error}`, 'scanForDataMashup', 'error');
			results.push({
				file: fileName,
				hasDataMashup: false,
				size: 0,
				error: String(error)
			});
		}
	}
	
	return results;
}

async function rawExtraction(uri?: vscode.Uri, uris?: vscode.Uri[]): Promise<void> {
	try {
		// Dump extension settings for debugging (debug level only)
		const logLevel = getConfig().get<string>('log.level', 'info');
		if (logLevel === 'debug') {
			dumpAllExtensionSettings();
		}
		
		// Handle multiple file selection (batch operations)
		if (uris && uris.length > 1) {
			log(`Batch raw extraction started: ${uris.length} Excel files selected`, 'rawExtraction', 'info');
			vscode.window.showInformationMessage(`Running raw extraction on ${uris.length} Excel files...`);
			
			let successCount = 0;
			let errorCount = 0;
			
			for (const fileUri of uris) {
				try {
					await rawExtraction(fileUri); // Recursive call for single file
					successCount++;
				} catch (error) {
					log(`Failed raw extraction from ${path.basename(fileUri.fsPath)}: ${error}`, 'rawExtraction', 'error');
					errorCount++;
				}
			}
			
			const resultMsg = `Batch raw extraction completed: ${successCount} successful, ${errorCount} failed`;
			log(resultMsg, 'rawExtraction', 'success');
			vscode.window.showInformationMessage(resultMsg);
			return;
		}
		
		// Validate URI parameter - don't show file dialog for invalid input
		if (uri && (!uri.fsPath || typeof uri.fsPath !== 'string')) {
			const errorMsg = 'Invalid URI parameter provided to rawExtraction command';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'rawExtraction', 'error');
			return;
		}
		
		// NEVER show file dialogs - extension works only through VS Code UI
		if (!uri?.fsPath) {
			const errorMsg = 'No Excel file specified. Use right-click on an Excel file or Command Palette with file open.';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'rawExtraction', 'error');
			return;
		}
		
		const excelFile = uri.fsPath;
		if (!excelFile) {
			return;
		}

		log(`Starting enhanced raw extraction for: ${path.basename(excelFile)}`, 'rawExtraction', 'info');

		// Create debug output directory (delete if exists)
		const baseName = path.basename(excelFile, path.extname(excelFile));
		const outputDir = path.join(path.dirname(excelFile), `${baseName}_debug_extraction`);
		
		// Clean up existing debug directory
		if (fs.existsSync(outputDir)) {
			log(`Cleaning up existing debug directory: ${outputDir}`, 'rawExtraction', 'info');
			fs.rmSync(outputDir, { recursive: true, force: true });
		}
		fs.mkdirSync(outputDir);
		log(`Created fresh debug directory: ${outputDir}`, 'rawExtraction', 'info');

		// Get file stats
		const fileStats = fs.statSync(excelFile);
		const fileSizeMB = (fileStats.size / (1024 * 1024)).toFixed(2);
		log(`File size: ${fileSizeMB} MB`, 'rawExtraction', 'debug');

		// Use JSZip to extract and examine the Excel file structure
		try {
			const JSZip = (await import('jszip')).default;
			log('Reading Excel file buffer...', 'rawExtraction', 'debug');
			const buffer = fs.readFileSync(excelFile);
			
			log('Loading ZIP structure...', 'rawExtraction', 'debug');
			const startTime = Date.now();
			const zip = await JSZip.loadAsync(buffer);
			const loadTime = Date.now() - startTime;
			log(`ZIP loaded in ${loadTime}ms`, 'rawExtraction', 'info');
			
			// List all files
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			log(`Found ${allFiles.length} files in ZIP structure`, 'rawExtraction', 'info');
			
			// Categorize files
			const customXmlFiles = allFiles.filter(f => f.startsWith('customXml/'));
			const xlFiles = allFiles.filter(f => f.startsWith('xl/'));
			const queryFiles = allFiles.filter(f => f.includes('quer') || f.includes('Query'));
			const connectionFiles = allFiles.filter(f => f.includes('connection'));
			
			log(`Files breakdown: ${customXmlFiles.length} customXml, ${xlFiles.length} xl/, ${queryFiles.length} query-related, ${connectionFiles.length} connection-related`, 'rawExtraction', 'info');

			// Enhanced DataMashup detection - use the same logic as main extraction
			const xmlFiles = allFiles.filter(f => f.toLowerCase().endsWith('.xml'));
			log(`Scanning ${xmlFiles.length} XML files for DataMashup content...`, 'rawExtraction', 'info');
			
			// Use the unified DataMashup detection function
			const dataMashupResults = await scanForDataMashup(zip, allFiles, outputDir, true);
			
			// Count DataMashup findings
			const dataMashupFiles = dataMashupResults.filter(r => r.hasDataMashup);
			const totalDataMashupSize = dataMashupFiles.reduce((sum, r) => sum + r.size, 0);
			
			log(`DataMashup scan complete: Found ${dataMashupFiles.length} files containing DataMashup (${(totalDataMashupSize / 1024).toFixed(1)} KB total)`, 'rawExtraction', 'info');

			// Create comprehensive debug report
			const debugInfo = {
				extractionReport: {
					file: excelFile,
					fileSize: `${fileSizeMB} MB`,
					extractedAt: new Date().toISOString(),
					zipLoadTime: `${loadTime}ms`,
					totalFiles: allFiles.length
				},
				fileStructure: {
					allFiles: allFiles,
					customXmlFiles: customXmlFiles,
					xlFiles: xlFiles,
					queryFiles: queryFiles,
					connectionFiles: connectionFiles
				},
				dataMashupAnalysis: {
					totalXmlFilesScanned: dataMashupResults.length,
					dataMashupFilesFound: dataMashupFiles.length,
					totalDataMashupSize: `${(totalDataMashupSize / 1024).toFixed(1)} KB`,
					results: dataMashupResults.map(r => ({
						file: r.file,
						hasDataMashup: r.hasDataMashup,
						size: r.size,
						...(r.error && { error: r.error }),
						...(r.extractedFormula && { 
							extractedFormulaSize: `${(r.extractedFormula.length / 1024).toFixed(1)} KB`,
							formulaPreview: r.extractedFormula.substring(0, 200) + '...'
						})
					}))
				},
				potentialPowerQueryLocations: customXmlFiles.concat([
					'xl/queryTables/queryTable1.xml',
					'xl/connections.xml'
				]).filter(loc => allFiles.includes(loc)),
				recommendations: dataMashupFiles.length === 0 ? 
					['No DataMashup content found - file may not contain Power Query M code', 'Check if Excel file actually has Power Query connections'] :
					[
						`Found DataMashup in: ${dataMashupFiles.map((f: DataMashupScanResult) => f.file).join(', ')}`, 
						'Use extracted DataMashup files for further analysis',
						...(dataMashupFiles.some((f: DataMashupScanResult) => f.extractedFormula) ? ['Successfully extracted M code - check _PowerQuery.m files'] : [])
					]
			};

			const reportPath = path.join(outputDir, 'EXTRACTION_REPORT.json');
			fs.writeFileSync(reportPath, JSON.stringify(debugInfo, null, 2), 'utf8');
			log(`Comprehensive report saved: ${path.basename(reportPath)}`, 'rawExtraction', 'info');

			// Show results
			const extractedCodeFiles = dataMashupFiles.filter((f: DataMashupScanResult) => f.extractedFormula).length;
			const message = dataMashupFiles.length > 0 ?
				`✅ Enhanced extraction completed!\n🔍 Found ${dataMashupFiles.length} DataMashup source(s) in ${path.basename(excelFile)}\n📁 Extracted ${extractedCodeFiles} M code file(s)\n📁 Results in: ${path.basename(outputDir)}` :
				`⚠️ Enhanced extraction completed!\n❌ No DataMashup content found in ${path.basename(excelFile)}\n📁 Debug files in: ${path.basename(outputDir)}`;
			
			vscode.window.showInformationMessage(message);
			log(message.replace(/\n/g, ' | '), 'rawExtraction', 'info');
			
		} catch (error) {
			log(`ZIP extraction/analysis failed: ${error}`, 'rawExtraction', 'info');

			// Write error info
			const debugInfo = {
				extractionReport: {
					file: excelFile,
					fileSize: `${fileSizeMB} MB`,
					extractedAt: new Date().toISOString(),
					error: 'Failed to extract Excel file structure',
					errorDetails: String(error)
				}
			};

			fs.writeFileSync(
				path.join(outputDir, 'ERROR_REPORT.json'),
				JSON.stringify(debugInfo, null, 2),
				'utf8'
			);
		}
		
	} catch (error) {
		const errorMsg = `Raw extraction failed: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, 'rawExtraction', 'debug');
		log(`Raw extraction error: ${error}`, 'rawExtraction', 'error');
	}
}

// New function to dump all extension settings for debugging
function dumpAllExtensionSettings(): void {
	try {
		log('=== EXTENSION SETTINGS DUMP ===', 'dumpAllExtensionSettings', 'debug');
		const extensionId = 'excel-power-query-editor';
		// Get all configuration scopes
		const userConfig = vscode.workspace.getConfiguration(extensionId, null);
		const workspaceConfig = vscode.workspace.getConfiguration(extensionId, vscode.workspace.workspaceFolders?.[0]?.uri);
		// Collect all keys from both configs
		const allKeys = new Set<string>();
		for (const key of Object.keys(userConfig)) { allKeys.add(key); }
		for (const key of Object.keys(workspaceConfig)) { allKeys.add(key); }
		// Always include logLevel
		// allKeys.add('log.level');
		// Dump each setting with its value and source
		for (const key of Array.from(allKeys).sort()) {
			let value: any = undefined;
			let source: string = 'default';
			if (workspaceConfig.has(key)) {
				value = workspaceConfig.get(key);
				source = 'workspace';
			} else if (userConfig.has(key)) {
				value = userConfig.get(key);
				source = 'user';
			} else {
				value = vscode.workspace.getConfiguration(extensionId).inspect(key)?.defaultValue;
			}
			log(`  ${key}: ${JSON.stringify(value)} [${source}]`, 'dumpAllExtensionSettings', 'debug');
		}
		// Check environment info
		log('ENVIRONMENT INFO:', 'dumpAllExtensionSettings', 'debug');
		log(`  Remote Name: ${vscode.env.remoteName || '<not remote>'}`, 'dumpAllExtensionSettings', 'info');
		log(`  VS Code Version: ${vscode.version}`, 'dumpAllExtensionSettings', 'info');
		log(`  Workspace Folders: ${vscode.workspace.workspaceFolders?.length || 0}`, 'dumpAllExtensionSettings', 'info');
		// Check if we're in a dev container
		const isDevContainer = vscode.env.remoteName?.includes('dev-container');
		log(`  Is Dev Container: ${isDevContainer}`, 'dumpAllExtensionSettings', 'info');
		log('=== END SETTINGS DUMP ===', 'dumpAllExtensionSettings', 'info');
	} catch (error) {
		log(`Failed to dump settings: ${error}`, 'dumpAllExtensionSettings', 'error');
	}
}
/**
 * Settings migration: v0.5.0 flat names -> namespaced names.
 *
 * WHAT THIS REPLACES. The previous implementation enumerated the configuration object and set
 * every key to `undefined` in both User and Workspace scope - a wipe, not a migration. It
 * preserved nothing, and its guard compared a stored marker against the EXTENSION VERSION, so it
 * re-ran on every release. Shipping that would have deleted the settings of every user of this
 * extension, twice.
 *
 * WHY THE OBVIOUS APPROACH DOES NOT WORK. VS Code only lets an extension write a configuration key
 * that is REGISTERED in `contributes.configuration`. The 0.5.x refactor renamed the settings and
 * deleted the old names from package.json in the same change, which removed the only handle a
 * migration has: `update(oldKey, undefined, ...)` on an unregistered key fails, so the stale value
 * stays in the user's settings.json forever with no API able to remove it.
 *
 * The old keys are therefore still DECLARED in package.json, marked deprecated. They render struck
 * through with a pointer to the replacement, and - the part that matters - they remain clearable.
 * Leave them in place for at least one minor release so a user who skips a version still migrates.
 *
 * RULES THIS FOLLOWS:
 *   - never write a scope the user did not already use - `inspect()` says where the value lives
 *   - never overwrite a new value the user has already set
 *   - only clear an old key in the scope it was actually set in
 *   - a failure on one key must not abandon the rest
 */

/** Bump this when the rename table changes. Deliberately NOT the extension version. */
const SETTINGS_MIGRATION_SCHEMA = '1';

/** Old key -> new key, relative to the `excel-power-query-editor` section. */
const LEGACY_RENAMES: ReadonlyArray<readonly [string, string]> = [
	['autoBackupBeforeSync', 'backup.autoBackupBeforeSync'],
	['autoCleanupBackups', 'backup.autoCleanup'],
	['backupLocation', 'backup.location'],
	['customBackupPath', 'backup.customPath'],
	['logLevel', 'log.level'],
	['showStatusBarInfo', 'log.showStatusBarInfo'],
	['syncDeleteAlwaysConfirm', 'sync.deleteAlwaysConfirm'],
	['syncTimeout', 'sync.timeout'],
	['watchAlways', 'watch.always'],
	['watchAlwaysMaxFiles', 'watch.maxFiles'],
	['watchOffOnDelete', 'watch.offOnDelete'],
];

type ScopeField = 'globalValue' | 'workspaceValue' | 'workspaceFolderValue';

/**
 * Move one scope's worth of settings.
 *
 * Returns both counts. `failed` matters as much as `moved`: the caller must not stamp a scope as
 * migrated when a key failed, or that key is skipped forever on every later activation and the
 * user's value is stranded under a name nothing reads.
 */
async function migrateScope(
	config: vscode.WorkspaceConfiguration,
	target: vscode.ConfigurationTarget,
	field: ScopeField
): Promise<{ moved: number; failed: number }> {
	let moved = 0;
	let failed = 0;

	for (const [oldKey, newKey] of LEGACY_RENAMES) {
		const oldInfo = config.inspect(oldKey);
		const oldValue = oldInfo ? (oldInfo as Record<string, unknown>)[field] : undefined;
		if (oldValue === undefined) { continue; }

		try {
			const newInfo = config.inspect(newKey);
			const newValue = newInfo ? (newInfo as Record<string, unknown>)[field] : undefined;

			// Only carry the value across if the user has not already set the new one here.
			if (newValue === undefined) {
				await config.update(newKey, oldValue, target);
				moved++;
				log(`Migrated ${oldKey} -> ${newKey} (${field})`, 'settingsMigration', 'info');
			} else {
				log(`Kept existing ${newKey} (${field}); dropped stale ${oldKey}`, 'settingsMigration', 'info');
			}

			await config.update(oldKey, undefined, target);
		} catch (error) {
			// One bad key must not strand the others - but it must be remembered, so this scope is
			// left unmarked and tried again next time.
			failed++;
			log(`Could not migrate ${oldKey} (${field}): ${error}`, 'settingsMigration', 'warn');
		}
	}

	const flags = await migrateLogFlags(config, target, field);
	return { moved: moved + flags.moved, failed: failed + flags.failed };
}

/**
 * debugMode and verboseMode were booleans; log.level is an enum. Precedence: an explicit
 * log.level the user already set wins, then debug, then verbose.
 */
async function migrateLogFlags(
	config: vscode.WorkspaceConfiguration,
	target: vscode.ConfigurationTarget,
	field: ScopeField
): Promise<{ moved: number; failed: number }> {
	const read = (key: string): unknown => {
		const info = config.inspect(key);
		return info ? (info as Record<string, unknown>)[field] : undefined;
	};

	const debugMode = read('debugMode');
	const verboseMode = read('verboseMode');
	if (debugMode === undefined && verboseMode === undefined) { return { moved: 0, failed: 0 }; }

	let moved = 0;
	let failed = 0;
	try {
		if (read('log.level') === undefined) {
			const level = debugMode === true ? 'debug' : verboseMode === true ? 'verbose' : undefined;
			if (level) {
				await config.update('log.level', level, target);
				moved++;
				log(`Migrated ${debugMode === true ? 'debugMode' : 'verboseMode'} -> log.level='${level}' (${field})`,
					'settingsMigration', 'info');
			}
		}
		if (debugMode !== undefined) { await config.update('debugMode', undefined, target); }
		if (verboseMode !== undefined) { await config.update('verboseMode', undefined, target); }
	} catch (error) {
		failed++;
		log(`Could not migrate log flags (${field}): ${error}`, 'settingsMigration', 'warn');
	}
	return { moved, failed };
}

/**
 * Which scopes still need migrating, given the migration marker's PER-SCOPE values.
 *
 * This has to be per scope, and getting it wrong is expensive. The marker is written to Global, but
 * `get()` returns the EFFECTIVE value - global merged with workspace and folder. A guard reading
 * `get('xtn.level')` therefore says "already migrated" in EVERY workspace as soon as the first one
 * is done, so a second project with legacy settings in its own `.vscode/settings.json` is skipped
 * forever. Nothing reads the old keys as a fallback, so those settings do not keep working - they
 * silently stop applying and the defaults take over. Someone who set `backup.location` to
 * `tempFolder` starts writing backups next to a workbook in their synced OneDrive folder instead.
 *
 * Exported so the decision table can be tested directly: the extension test host runs with no
 * workspace folder open, so the workspace and folder paths cannot be exercised end to end.
 */
export function scopesNeedingMigration(
	marker: { globalValue?: unknown; workspaceValue?: unknown; workspaceFolderValue?: unknown } | undefined,
	schema: string
): { global: boolean; workspace: boolean; folder: boolean } {
	return {
		global: marker?.globalValue !== schema,
		workspace: marker?.workspaceValue !== schema,
		folder: marker?.workspaceFolderValue !== schema
	};
}

export async function migrateLegacySettings(): Promise<void> {
	const extensionId = 'excel-power-query-editor';
	const MARKER = 'xtn.level';

	try {
		const root = vscode.workspace.getConfiguration(extensionId);
		const needed = scopesNeedingMigration(
			root.inspect(MARKER) as Record<string, unknown> | undefined,
			SETTINGS_MIGRATION_SCHEMA
		);

		let moved = 0;

		if (needed.global) {
			const result = await migrateScope(root, vscode.ConfigurationTarget.Global, 'globalValue');
			moved += result.moved;
			// Mark only a scope that fully succeeded. Stamping it after a per-key failure means that
			// key is never retried, and the user's value stays under a name nothing reads.
			if (result.failed === 0) {
				await root.update(MARKER, SETTINGS_MIGRATION_SCHEMA, vscode.ConfigurationTarget.Global);
			}
		}

		if (needed.workspace && vscode.workspace.workspaceFolders?.length) {
			const result = await migrateScope(root, vscode.ConfigurationTarget.Workspace, 'workspaceValue');
			moved += result.moved;
			// Only mark a workspace we actually changed, and only if nothing failed. Writing a marker
			// into someone's committed .vscode/settings.json for no reason puts our bookkeeping in
			// their next git diff, and re-scanning a clean workspace costs a few inspect() calls.
			if (result.moved > 0 && result.failed === 0) {
				await root.update(MARKER, SETTINGS_MIGRATION_SCHEMA, vscode.ConfigurationTarget.Workspace);
			}
		}

		// Folder scope needs a folder-scoped configuration object to read and write through, and its
		// own marker per folder - a multi-root workspace can have legacy settings in any of them.
		for (const folder of vscode.workspace.workspaceFolders ?? []) {
			const folderConfig = vscode.workspace.getConfiguration(extensionId, folder.uri);
			const folderMarker = folderConfig.inspect(MARKER) as Record<string, unknown> | undefined;
			if (folderMarker?.workspaceFolderValue === SETTINGS_MIGRATION_SCHEMA) { continue; }

			const result = await migrateScope(
				folderConfig, vscode.ConfigurationTarget.WorkspaceFolder, 'workspaceFolderValue');
			moved += result.moved;
			if (result.moved > 0 && result.failed === 0) {
				await folderConfig.update(MARKER, SETTINGS_MIGRATION_SCHEMA, vscode.ConfigurationTarget.WorkspaceFolder);
			}
		}

		log(`Settings migration complete - ${moved} value(s) carried across (schema ${SETTINGS_MIGRATION_SCHEMA})`,
			'settingsMigration', moved > 0 ? 'info' : 'debug');
	} catch (error) {
		// Never let a settings migration stop the extension activating.
		log(`Settings migration failed: ${error}`, 'settingsMigration', 'error');
	}
}

async function findExcelFile(mFilePath: string): Promise<string | undefined> {
	const dir = path.dirname(mFilePath);
	const mFileName = path.basename(mFilePath, '.m');
	
	// Remove '_PowerQuery' suffix to get original Excel filename
	if (mFileName.endsWith('_PowerQuery')) {
		const originalFileName = mFileName.replace(/_PowerQuery$/, '');
		const candidatePath = path.join(dir, originalFileName);
		
		if (fs.existsSync(candidatePath)) {
			return candidatePath;
		}
	}
	
	return undefined;
}

async function cleanupBackupsCommand(uri?: vscode.Uri): Promise<void> {
	try {
		// Migrate legacy settings on every activation
		await migrateLegacySettings();
		// Validate URI parameter - don't show file dialog for invalid input
		if (uri && (!uri.fsPath || typeof uri.fsPath !== 'string')) {
			const errorMsg = 'Invalid URI parameter provided to cleanupBackups command';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'cleanupBackupsCommand', 'error');
			return;
		}
		
		// NEVER show file dialogs - extension works only through VS Code UI
		if (!uri?.fsPath) {
			const errorMsg = 'No Excel file specified. Use right-click on an Excel file or Command Palette with file open.';
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, 'cleanupBackupsCommand', 'error');
			return;
		}
		
		const excelFile = uri.fsPath;

		const config = getConfig();
		const maxBackups = config.get<number>('backup.maxFiles', 5) || 5;
		
		// Get backup information
		const sampleTimestamp = '2000-01-01T00-00-00-000Z';
		const sampleBackupPath = getBackupPath(excelFile, sampleTimestamp);
		const backupDir = path.dirname(sampleBackupPath);
		const baseFileName = path.basename(excelFile);
		
		if (!fs.existsSync(backupDir)) {
			vscode.window.showInformationMessage(`No backup directory found for ${path.basename(excelFile)}`);
			return;
		}
		
		// Count existing backups
		const backupPattern = `${baseFileName}.backup.`;
		const allFiles = fs.readdirSync(backupDir);
		const backupFiles = allFiles.filter(file => file.startsWith(backupPattern));
		
		if (backupFiles.length === 0) {
			vscode.window.showInformationMessage(`No backup files found for ${path.basename(excelFile)}`);
			return;
		}
		
		const willKeep = Math.min(backupFiles.length, maxBackups);
		const willDelete = Math.max(0, backupFiles.length - maxBackups);
		
		if (willDelete === 0) {
			vscode.window.showInformationMessage(`${backupFiles.length} backup files found for ${path.basename(excelFile)}. All within limit of ${maxBackups}.`);
			return;
		}
		
		const confirmation = await vscode.window.showWarningMessage(
			`Found ${backupFiles.length} backup files for ${path.basename(excelFile)}.\n` +
			`Keep ${willKeep} most recent, delete ${willDelete} oldest?`,
			{ modal: true },
			'Yes, Cleanup', 'Cancel'
		);
		
		if (confirmation === 'Yes, Cleanup') {
			// Force cleanup by temporarily enabling auto-cleanup
			const originalAutoCleanup = config.get<boolean>('backup.autoCleanup', true);
			if (config.update) {
				await config.update('backup.autoCleanup', true, vscode.ConfigurationTarget.Global);
			}
			
			try {
				cleanupOldBackups(excelFile);
				vscode.window.showInformationMessage(`✅ Backup cleanup completed for ${path.basename(excelFile)}`);
			} finally {
				// Restore original setting
				if (config.update) {
					await config.update('backup.autoCleanup', originalAutoCleanup, vscode.ConfigurationTarget.Global);
				}
			}
		}
		
	} catch (error) {
		const errorMsg = `Failed to cleanup backups: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(`Backup cleanup error: ${error}`, 'cleanupBackupsCommand', 'error');
	}
}

// Install Excel Power Query symbols for IntelliSense
async function installExcelSymbols(): Promise<void> {
	// Now a report rather than an install. The symbols are registered at activation through the
	// Power Query extension's API; this command exists so a user can ask whether it worked, and
	// retry if they installed the PQ extension after us.
	const extensionPath = vscode.extensions
		.getExtension('ewc3labs.excel-power-query-editor')?.extensionPath;

	if (!extensionPath) {
		vscode.window.showErrorMessage('Could not locate the extension directory.');
		return;
	}

	const result = await registerExcelSymbols(extensionPath, (m, l) => log(m, 'excelSymbols', l ?? 'info'));
	const message = explainRegistration(result);
	log(message, 'excelSymbols', result.ok ? 'success' : 'warn');

	if (result.ok) {
		vscode.window.showInformationMessage(message);
	} else {
		vscode.window.showWarningMessage(message);
	}
}


// Debounced sync helper to prevent multiple syncs in rapid succession
async function debouncedSyncToExcel(mFile: string): Promise<void> {
	// Check if this file was recently extracted - if so, skip auto-sync
	if (recentExtractions.has(mFile)) {
		log(`Skipping auto-sync for recently extracted file: ${path.basename(mFile)}`, 'debouncedSyncToExcel', 'verbose');
		return;
	}
	
	const config = getConfig();
	let debounceMs = config.get<number>('sync.debounceMs', 500) || 500;
	
	// Get Excel file size to determine appropriate debounce timing
	let fileSize = 0;
	try {
		// Find the corresponding Excel file to check its size
		const excelFile = await findExcelFile(mFile);
		if (excelFile && fs.existsSync(excelFile)) {
			const stats = fs.statSync(excelFile);
			fileSize = stats.size;
		}
	} catch (error) {
		// If we can't get Excel file size, use default debounce
	}
	
	// Apply intelligent debouncing based on Excel file size
	const fileSizeMB = fileSize / (1024 * 1024);
	const largeFileMinDebounce = config.get<number>('sync.largefile.minDebounceMs', 5000) || 5000;
	
	if (fileSizeMB > 50) {
		// For files over 50MB, use configurable minimum debounce (default 5 seconds)
		debounceMs = Math.max(debounceMs, largeFileMinDebounce);
		log(`Large file detected (${fileSizeMB.toFixed(1)}MB), using extended debounce: ${debounceMs}ms`, 'debouncedSyncToExcel', 'verbose');
	} else if (fileSizeMB > 10) {
		// For files over 10MB, use half the large file debounce
		const mediumFileDebounce = Math.max(2000, largeFileMinDebounce / 2);
		debounceMs = Math.max(debounceMs, mediumFileDebounce);
		log(`Medium file detected (${fileSizeMB.toFixed(1)}MB), using extended debounce: ${debounceMs}ms`, 'debouncedSyncToExcel', 'verbose');
	}
	
	// Only execute immediately if debounce is explicitly set to 0 (not just small)
	if (debounceMs === 0) {
		log(`IMMEDIATE SYNC (debounce explicitly disabled) for ${path.basename(mFile)}`, 'debouncedSyncToExcel', 'verbose');
		syncToExcel(vscode.Uri.file(mFile)).catch(error => {
			log(`Immediate sync failed for ${path.basename(mFile)}: ${error}`, 'debouncedSyncToExcel', 'error');
		});
		return;
	}
	
	// Clear existing timer for this file
	const existingTimer = debounceTimers.get(mFile);
	if (existingTimer) {
		clearTimeout(existingTimer);
	}
	
	// Set new timer
	const timer = setTimeout(async () => {
		try {
			log(`Debounced sync executing for ${path.basename(mFile)}`, 'debouncedSyncToExcel', 'verbose');
			await syncToExcel(vscode.Uri.file(mFile));
			debounceTimers.delete(mFile);
		} catch (error) {
			log(`Debounced sync failed for ${path.basename(mFile)}: ${error}`, 'debouncedSyncToExcel', 'error');
			debounceTimers.delete(mFile);
		}
	}, debounceMs);
	
	debounceTimers.set(mFile, timer);
	log(`Sync debounced for ${path.basename(mFile)} (${debounceMs}ms)`, 'debouncedSyncToExcel', 'verbose');
}

// Check if Excel file is writable (not locked)
async function isExcelFileWritable(excelFile: string): Promise<boolean> {
	const config = getConfig();
	const checkWriteable = config.get<boolean>('watch.checkExcelWriteable', true);
	
	if (!checkWriteable) {
		return true; // Skip check if disabled
	}
	
	try {
		// Try to open the file for writing to check if it's locked
		const handle = await fs.promises.open(excelFile, 'r+');
		await handle.close();
		return true;
	} catch (error: any) {
		// File is likely locked by Excel or another process
		log(`Excel file appears to be locked: ${error.message}`, 'isExcelFileWritable', 'debug');
		return false;
	}
}

// This method is called when your extension is deactivated

/**
 * Report - never remove - what the pre-PQ-18 file-based symbols version left behind.
 *
 * Shown at most once per machine. A notification that returns every startup is one people learn to
 * dismiss without reading, which would defeat the point on the one startup it mattered.
 */
const LEGACY_SYMBOLS_NOTICE_KEY = 'excelPowerQueryEditor.legacySymbolsNoticeShown';

async function reportLegacySymbolLeftovers(context: vscode.ExtensionContext): Promise<void> {
	try {
		if (context.globalState.get<boolean>(LEGACY_SYMBOLS_NOTICE_KEY)) { return; }

		const { folders, settingStillPoints } = findLegacyLeftovers();
		if (folders.length === 0 && !settingStillPoints) { return; }

		log(`Legacy symbols leftovers found: ${folders.length} folder(s), `
			+ `setting still points: ${settingStillPoints}`, 'excelSymbols', 'info');

		// Two different situations, and only one of them is actively doing something.
		const message = settingStillPoints
			? "Excel symbols now come from the Power Query API, but this extension's old symbols "
				+ "folder is still listed in powerquery.client.additionalSymbolsDirectories - so a "
				+ "stale copy is still being loaded. Nothing has been deleted."
			: "This extension used to write an Excel symbols file to disk. It no longer does, and "
				+ "the old folder is still there. Nothing has been deleted.";

		const SHOW = 'Show Me';
		const DISMISS = "Don't Show Again";
		const choice = await vscode.window.showInformationMessage(message, SHOW, DISMISS);

		if (choice === SHOW && folders.length > 0) {
			await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(folders[0]));
		}
		if (choice === SHOW || choice === DISMISS) {
			await context.globalState.update(LEGACY_SYMBOLS_NOTICE_KEY, true);
		}
	} catch (e) {
		// Reporting leftovers must never be the thing that breaks a session.
		log(`Could not check for legacy symbols leftovers: ${e instanceof Error ? e.message : e}`,
			'excelSymbols', 'debug');
	}
}

export function deactivate() {
	// Take our symbols back out, so the language service is left as we found it.
	void unregisterExcelSymbols();

	// Close all file watchers
	for (const [, watchers] of fileWatchers) {
		watchers.chokidar.close();
		watchers.vscode?.dispose();
		watchers.document?.dispose();
	}
	fileWatchers.clear();
}

// Parse structured metadata from .m file header
