// The module 'vscode' contains the VS Code extensibility API
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { watch, FSWatcher } from 'chokidar';
import { getConfig } from './configHelper';

// Test environment detection
function isTestEnvironment(): boolean {
	return process.env.NODE_ENV === 'test' || 
	       process.env.VSCODE_TEST_ENV === 'true' ||
	       typeof global.describe !== 'undefined'; // Jest/Mocha detection
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
let outputChannel: vscode.OutputChannel;

// Status bar item for watch status
let statusBarItem: vscode.StatusBarItem;

// Backup path helper
function getBackupPath(excelFile: string, timestamp: string): string {
	const config = getConfig();
	const backupLocation = config.get<string>('backupLocation', 'sameFolder');
	const baseFileName = path.basename(excelFile);
	const backupFileName = `${baseFileName}.backup.${timestamp}`;
	
	switch (backupLocation) {
		case 'tempFolder':
			return path.join(require('os').tmpdir(), 'excel-pq-backups', backupFileName);
		case 'custom':
			const customPath = config.get<string>('customBackupPath', '');
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
	const autoCleanup = config.get<boolean>('autoCleanupBackups', true) || false;
	
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
					log(`Deleted old backup: ${backup.filename}`);
				} catch (deleteError) {
					log(`Failed to delete backup ${backup.filename}: ${deleteError}`, 'cleanupBackups');
				}
			}
			
			if (deletedCount > 0) {
				log(`Cleaned up ${deletedCount} old backup files (keeping ${maxBackups} most recent)`);
			}
		}
		
	} catch (error) {
		log(`Backup cleanup failed: ${error}`, 'cleanupBackups');
	}
}

// Enhanced logging function with context and log levels
function log(message: string, context?: string): void {
	const config = getConfig();
	const logLevel = getEffectiveLogLevel();
	
	// Determine message level based on context or content
	let messageLevel = 'info';
	if (context === 'error' || message.includes('❌') || message.toLowerCase().includes('error')) {
		messageLevel = 'error';
	} else if (message.includes('⚠️') || message.toLowerCase().includes('warning')) {
		messageLevel = 'warn';
	} else if (context === 'debug' || context === 'extractPowerQuery' || context === 'syncToExcel' || context === 'watchFile') {
		messageLevel = 'verbose';
	}
	
	// Check if message should be logged at current level
	const levelOrder = ['none', 'error', 'warn', 'info', 'verbose', 'debug'];
	const currentLevelIndex = levelOrder.indexOf(logLevel);
	const messageLevelIndex = levelOrder.indexOf(messageLevel);
	
	if (currentLevelIndex < messageLevelIndex) {
		return; // Don't log this message at current level
	}
	
	const timestamp = new Date().toISOString();
	const contextInfo = context ? `[${context}] ` : '';
	const fullMessage = `[${timestamp}] ${contextInfo}${message}`;
	console.log(fullMessage);
	
	// Only append to output channel if it's initialized
	if (outputChannel) {
		outputChannel.appendLine(fullMessage);
	}
}

// Get effective log level with automatic migration from legacy settings
function getEffectiveLogLevel(): string {
	const config = getConfig();
	
	// Check if new setting exists
	const logLevel = config.get<string>('logLevel');
	if (logLevel) {
		return logLevel;
	}
	
	// Check legacy settings and migrate
	const verboseMode = config.get<boolean>('verboseMode');
	const debugMode = config.get<boolean>('debugMode');
	
	let migratedLevel = 'info'; // Default
	
	if (debugMode) {
		migratedLevel = 'debug';
	} else if (verboseMode) {
		migratedLevel = 'verbose';
	}
	
	// Perform one-time migration if legacy settings exist
	if (verboseMode !== undefined || debugMode !== undefined) {
		// Use unified config system for migration
		const unifiedConfig = getConfig();
		if (unifiedConfig.update) {
			// Use Promise for async operation
			Promise.resolve(unifiedConfig.update('logLevel', migratedLevel, vscode.ConfigurationTarget.Global))
				.then(() => {
					vscode.window.showInformationMessage(
						`Excel Power Query Editor: Updated logging settings. ` +
						`Your previous settings (verbose: ${verboseMode}, debug: ${debugMode}) ` +
						`have been migrated to logLevel: "${migratedLevel}". ` +
						`Legacy settings can be safely removed from your configuration.`,
						'OK', 'Open Settings'
					).then(choice => {
						if (choice === 'Open Settings') {
							vscode.commands.executeCommand('workbench.action.openSettings', 'excel-power-query-editor');
						}
					});
					log(`Migrated legacy logging settings to logLevel: ${migratedLevel}`, 'migration');
				})
				.catch((error: any) => {
					log(`Failed to migrate legacy settings: ${error}`, 'error');
				});
		} else {
			// Test environment - just log the migration intent
			log(`Test environment: Would migrate legacy logging settings to logLevel: ${migratedLevel}`, 'migration');
		}
	}
	
	return migratedLevel;
}

// Update status bar
function updateStatusBar() {
	const config = getConfig();
	if (!config.get<boolean>('showStatusBarInfo', true)) {
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
	const watchAlways = config.get<boolean>('watchAlways', false);
	
	if (!watchAlways) {
		log('Extension activated - auto-watch disabled, staying dormant until manual command');
		return; // Auto-watch is disabled - minimal initialization
	}

	log('Extension activated - auto-watch enabled, scanning workspace for .m files...');

	try {
		// Find all .m files in the workspace
		const mFiles = await vscode.workspace.findFiles('**/*.m', '**/node_modules/**');
		
		if (mFiles.length === 0) {
			log('Auto-watch enabled but no .m files found in workspace');
			vscode.window.showInformationMessage('🔍 Auto-watch enabled but no .m files found in workspace');
			return;
		}

		log(`Found ${mFiles.length} .m files in workspace, checking for corresponding Excel files...`);

		let watchedCount = 0;
		const maxAutoWatch = 20; // Prevent watching too many files automatically
		
		for (const mFileUri of mFiles.slice(0, maxAutoWatch)) {
			const mFile = mFileUri.fsPath;
			
			// Check if there's a corresponding Excel file
			const excelFile = await findExcelFile(mFile);
			if (excelFile && fs.existsSync(excelFile)) {
				try {
					await watchFile(mFileUri);
					watchedCount++;
					log(`Auto-watch initialized: ${path.basename(mFile)} → ${path.basename(excelFile)}`);
				} catch (error) {
					log(`Failed to auto-watch ${path.basename(mFile)}: ${error}`, 'autoWatchInit');
				}
			} else {
				log(`Skipping ${path.basename(mFile)} - no corresponding Excel file found`);
			}
		}

		if (watchedCount > 0) {
			vscode.window.showInformationMessage(
				`🚀 Auto-watch enabled: Now watching ${watchedCount} Power Query file${watchedCount > 1 ? 's' : ''}`
			);
			log(`Auto-watch initialization complete: ${watchedCount} files being watched`);
		} else {
			log('Auto-watch enabled but no .m files with corresponding Excel files found');
			vscode.window.showInformationMessage('⚠️ Auto-watch enabled but no .m files with corresponding Excel files found');
		}

		if (mFiles.length > maxAutoWatch) {
			vscode.window.showWarningMessage(
				`Found ${mFiles.length} .m files but only auto-watching first ${maxAutoWatch}. Use "Watch File" command for others.`
			);
			log(`Limited auto-watch to ${maxAutoWatch} files (found ${mFiles.length} total)`);
		}

	} catch (error) {
		log(`Auto-watch initialization failed: ${error}`, 'autoWatchInit');
		vscode.window.showErrorMessage(`Auto-watch initialization failed: ${error}`);
	}
}

// This method is called when your extension is activated
export async function activate(context: vscode.ExtensionContext) {
	try {
		// Initialize output channel first (before any logging)
		outputChannel = vscode.window.createOutputChannel('Excel Power Query Editor');
		
		log('Excel Power Query Editor extension is now active!', 'activation');

		// Register all commands
		const commands = [
			vscode.commands.registerCommand('excel-power-query-editor.extractFromExcel', extractFromExcel),
			vscode.commands.registerCommand('excel-power-query-editor.syncToExcel', syncToExcel),
			vscode.commands.registerCommand('excel-power-query-editor.watchFile', watchFile),
			vscode.commands.registerCommand('excel-power-query-editor.toggleWatch', toggleWatch),
			vscode.commands.registerCommand('excel-power-query-editor.stopWatching', stopWatching),
			vscode.commands.registerCommand('excel-power-query-editor.syncAndDelete', syncAndDelete),
			vscode.commands.registerCommand('excel-power-query-editor.rawExtraction', rawExtraction),
			vscode.commands.registerCommand('excel-power-query-editor.cleanupBackups', cleanupBackupsCommand),
			vscode.commands.registerCommand('excel-power-query-editor.applyRecommendedDefaults', applyRecommendedDefaults)
		];

		context.subscriptions.push(...commands);
		log(`Registered ${commands.length} commands successfully`, 'activation');

		// Initialize status bar
		updateStatusBar();
		
		log('Excel Power Query Editor extension activated');
		
		// Auto-watch existing .m files if setting is enabled
		await initializeAutoWatch();
		
		log('Extension activation completed successfully', 'activation');
	} catch (error) {
		console.error('Extension activation failed:', error);
		// Re-throw to ensure VS Code knows about the failure
		throw error;
	}
}

async function extractFromExcel(uri?: vscode.Uri): Promise<void> {
	try {
		// Dump extension settings for debugging (debug level only)
		const logLevel = getEffectiveLogLevel();
		if (logLevel === 'debug') {
			dumpAllExtensionSettings();
		}
		
		const excelFile = uri?.fsPath || await selectExcelFile();
		if (!excelFile) {
			log('No Excel file selected for extraction');
			return;
		}

		log(`Starting Power Query extraction from: ${path.basename(excelFile)}`, 'extractPowerQuery');
		vscode.window.showInformationMessage(`Extracting Power Query from: ${path.basename(excelFile)}`);
		
		// Try to use excel-datamashup for extraction
		try {
			log('Loading required modules...', 'extractPowerQuery');
			// First, we need to extract the DataMashup XML from the Excel file (scanning all customXml files)
			const JSZip = (await import('jszip')).default;
			
			// Use require for excel-datamashup to avoid ES module issues
			const excelDataMashup = require('excel-datamashup');
			log('Modules loaded successfully', 'extractPowerQuery');
					log('Reading Excel file buffer...', 'extractPowerQuery');
		let buffer: Buffer;
		try {
			buffer = fs.readFileSync(excelFile);
			const fileSizeMB = (buffer.length / (1024 * 1024)).toFixed(2);
			log(`Excel file read: ${fileSizeMB} MB`);
		} catch (error) {
			const errorMsg = `Failed to read Excel file: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, "error");
			return;
		}

		log('Loading ZIP structure...');
		let zip: any;
		try {
			zip = await JSZip.loadAsync(buffer, {
				checkCRC32: false // Skip CRC check for better performance on large files
			});
			log('ZIP structure loaded successfully');
		} catch (error) {
			const errorMsg = `Failed to load Excel file as ZIP: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, "error");
			return;
		}
			
			// Debug: List all files in the Excel zip
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			log(`Files in Excel archive: ${allFiles.length} total files`, 'extractPowerQuery');
			
			// Look for Power Query DataMashup (the only format with actual M code)
			// Scan ALL customXml files instead of just hardcoded item1/2/3
			const customXmlFiles = Object.keys(zip.files)
				.filter(name => name.startsWith('customXml/') && name.endsWith('.xml'))
				.filter(name => !name.includes('/_rels/')) // Exclude relationship files
				.sort(); // Process in consistent order
			
			log(`Found ${customXmlFiles.length} customXml files to scan: ${customXmlFiles.join(', ')}`);
			
			let xmlContent: string | null = null;
			let foundLocation = '';
			
			for (const location of customXmlFiles) {
				const xmlFile = zip.file(location);
				if (xmlFile) {
					try {
						// Read as binary first, then decode properly
						const binaryData = await xmlFile.async('nodebuffer');
						let content: string;
						
						// Check for UTF-16 LE BOM (FF FE)
						if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
							log(`Detected UTF-16 LE BOM in ${location}`);
							// Decode UTF-16 LE (skip the 2-byte BOM)
							content = binaryData.subarray(2).toString('utf16le');
						} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
							log(`Detected UTF-8 BOM in ${location}`);
							// Decode UTF-8 (skip the 3-byte BOM)
							content = binaryData.subarray(3).toString('utf8');
						} else {
							// Try UTF-8 first (most common)
							content = binaryData.toString('utf8');
						}
						
						log(`Scanning ${location} for DataMashup content (${(content.length / 1024).toFixed(1)} KB)`);
						
						// Only accept DataMashup format - the only one with actual Power Query M code
						if (content.includes('DataMashup')) {
							xmlContent = content;
							foundLocation = location;
							log(`✅ Found DataMashup Power Query in: ${location}`);
							break; // Found actual Power Query, stop searching
						} else {
							log(`❌ No DataMashup content in ${location}`);
						}
					} catch (e) {
						log(`❌ Could not read ${location}: ${e}`);
					}
				}
			}
			
			if (!xmlContent) {
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
			
			log(`Attempting to parse DataMashup Power Query from: ${foundLocation}`);
			log(`DataMashup XML content size: ${(xmlContent.length / 1024).toFixed(2)} KB`);
			
			// Use excel-datamashup for DataMashup format
			log('Calling excelDataMashup.ParseXml()...');
			const parseResult = await excelDataMashup.ParseXml(xmlContent);
			log(`ParseXml() completed. Result type: ${typeof parseResult}`);
			
			if (typeof parseResult === 'string') {
				const errorMsg = `Power Query parsing failed: ${parseResult}\nLocation: ${foundLocation}\nXML preview: ${xmlContent.substring(0, 200)}...`;
				log(errorMsg, 'extraction');
				vscode.window.showErrorMessage(errorMsg);
				return;
			}
			
			log('ParseXml() succeeded. Extracting formula...');
			let formula: string;
			try {
				// Extract the formula
				formula = parseResult.getFormula();
				log(`getFormula() completed. Formula length: ${formula ? formula.length : 'null'}`);
			} catch (formulaError) {
				const errorMsg = `Formula extraction failed: ${formulaError}`;
				log(errorMsg, "error");
				vscode.window.showErrorMessage(errorMsg);
				return;
			}
			
			if (!formula) {
				const warningMsg = `No Power Query formula found in ${foundLocation}. ParseResult keys: ${Object.keys(parseResult).join(', ')}`;
				log(warningMsg, "error");
				vscode.window.showWarningMessage(warningMsg);
				return;
			}
			
			log('Formula extracted successfully. Creating output file...');
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
			
			vscode.window.showInformationMessage(`Power Query extracted to: ${path.basename(outputPath)}`);		log(`Successfully extracted Power Query from ${path.basename(excelFile)} to ${path.basename(outputPath)}`);
		
		// Track this file as recently extracted to prevent immediate auto-sync
		recentExtractions.add(outputPath);
		setTimeout(() => {
			recentExtractions.delete(outputPath);
			log(`Cleared recent extraction flag for ${path.basename(outputPath)}`, 'extractPowerQuery');
		}, 2000); // Prevent auto-sync for 2 seconds after extraction
		
		// Auto-watch if enabled
		const config = getConfig();
		if (config.get<boolean>('watchAlways', false)) {
			await watchFile(vscode.Uri.file(outputPath));
			log(`Auto-watch enabled for ${path.basename(outputPath)}`);
		}

		} catch (moduleError) {
			// Fallback: create a placeholder file
			const errorMsg = `Excel DataMashup parsing failed: ${moduleError}`;
			log(errorMsg, "error");
			log(`Error stack: ${moduleError instanceof Error ? moduleError.stack : 'No stack trace'}`);
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
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each [Column1] <> null),
    Result = #"Filtered Rows"
in
    Result`;

			fs.writeFileSync(outputPath, placeholderContent, 'utf8');
			
			// Open the created file
			const document = await vscode.workspace.openTextDocument(outputPath);
			await vscode.window.showTextDocument(document);
					vscode.window.showInformationMessage(`Placeholder file created: ${path.basename(outputPath)}`);
		log(`Created placeholder file: ${path.basename(outputPath)}`);
		
		// Track this file as recently extracted to prevent immediate auto-sync
		recentExtractions.add(outputPath);
		setTimeout(() => {
			recentExtractions.delete(outputPath);
			log(`Cleared recent extraction flag for placeholder ${path.basename(outputPath)}`, 'extractPowerQuery');
		}, 2000); // Prevent auto-sync for 2 seconds after extraction
		
		// Auto-watch if enabled
		const config = getConfig();
		if (config.get<boolean>('watchAlways', false)) {
			await watchFile(vscode.Uri.file(outputPath));
			log(`Auto-watch enabled for placeholder ${path.basename(outputPath)}`);
		}
		}
		
	} catch (error) {
		const errorMsg = `Failed to extract Power Query: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
		console.error('Extract error:', error);
	}
}

async function syncToExcel(uri?: vscode.Uri): Promise<void> {
	let backupPath: string | null = null;
	
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			vscode.window.showErrorMessage('Please select or open a .m file to sync.');
			return;
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
						log(`Test environment: Using fixture ${fixture} for sync`, 'syncToExcel');
						break;
					}
				}
				if (!excelFile) {
					log('Test environment: No Excel fixtures found, skipping sync', 'syncToExcel');
					return;
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
				log(`SAFETY STOP: Refusing to sync ${mFileName} - corresponding Excel file not found`, 'syncToExcel');
				return; // HARD STOP - no file picker
			}
		}

		// Check if Excel file is writable (not locked by Excel or another process)
		const isWritable = await isExcelFileWritable(excelFile);
		if (!isWritable) {
			const fileName = path.basename(excelFile);
			const retry = await vscode.window.showWarningMessage(
				`Excel file "${fileName}" appears to be locked (possibly open in Excel). Close the file and try again.`,
				'Retry', 'Cancel'
			);
			if (retry === 'Retry') {
				// Retry after a short delay
				setTimeout(() => syncToExcel(uri), 1000);
			}
			return;
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
		log(`Header stripping - Found section at position ${headerLength}, removed ${headerLength} header characters`, 'syncToExcel');
	} else {
		// No section found - use original content (might be a different format)
		cleanMCode = mContent.trim();
		log(`Header stripping - No section declaration found, using original content`, 'syncToExcel');
	}
		
		if (!cleanMCode) {
			vscode.window.showErrorMessage('No Power Query M code found in file.');
			return;
		}
		
		// Create backup of Excel file if enabled
		const config = getConfig();
		
		if (config.get<boolean>('autoBackupBeforeSync', true)) {
			const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
			backupPath = getBackupPath(excelFile, timestamp);
			
			// Ensure backup directory exists
			const backupDir = path.dirname(backupPath);
			if (!fs.existsSync(backupDir)) {
				fs.mkdirSync(backupDir, { recursive: true });
			}
			
			fs.copyFileSync(excelFile, backupPath);
			vscode.window.showInformationMessage(`Syncing to Excel... (Backup created: ${path.basename(backupPath)})`);
			log(`Backup created: ${backupPath}`);
			
			// Clean up old backups
			cleanupOldBackups(excelFile);
		} else {
			vscode.window.showInformationMessage(`Syncing to Excel... (No backup - disabled in settings)`);
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
		
		// Scan all customXml files for DataMashup content using same logic as extraction
		log('Scanning all customXml files for DataMashup content...', 'syncToExcel');
		for (const location of customXmlFiles) {
			const file = zip.file(location);
			if (file) {
				try {
					// Use same binary reading and BOM handling as extraction
					const binaryData = await file.async('nodebuffer');
					let content: string;
					
					// Check for UTF-16 LE BOM (FF FE)
					if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
						log(`Detected UTF-16 LE BOM in ${location}`, 'syncToExcel');
						content = binaryData.subarray(2).toString('utf16le');
					} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
						log(`Detected UTF-8 BOM in ${location}`, 'syncToExcel');
						content = binaryData.subarray(3).toString('utf8');
					} else {
						content = binaryData.toString('utf8');
					}
					
					if (content.includes('DataMashup')) {
						dataMashupFile = file;
						dataMashupLocation = location;
						log(`✅ Found DataMashup for sync in: ${location}`, 'syncToExcel');
						break;
					}
				} catch (e) {
					log(`Could not check ${location}: ${e}`, 'syncToExcel');
				}
			}
		}
		
		if (!dataMashupFile) {
			vscode.window.showErrorMessage('No DataMashup found in Excel file. This file may not contain Power Query.');
			return;
		}
		
		// Read and decode the DataMashup XML
		const binaryData = await dataMashupFile.async('nodebuffer');
		let dataMashupXml: string;
		
		// Handle UTF-16 LE BOM like in extraction
		if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
			log('Detected UTF-16 LE BOM in DataMashup', 'syncToExcel');
			dataMashupXml = binaryData.subarray(2).toString('utf16le');
		} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
			log('Detected UTF-8 BOM in DataMashup', 'syncToExcel');
			dataMashupXml = binaryData.subarray(3).toString('utf8');
		} else {
			dataMashupXml = binaryData.toString('utf8');
		}
		
		if (!dataMashupXml.includes('DataMashup')) {
			vscode.window.showErrorMessage('Invalid DataMashup format in Excel file.');
			return;
		}
		
		// DEBUG: Save the original DataMashup XML for inspection
		const debugDir = path.join(path.dirname(excelFile), 'debug_sync');
		if (!fs.existsSync(debugDir)) {
			fs.mkdirSync(debugDir);
		}
		fs.writeFileSync(
			path.join(debugDir, 'original_datamashup.xml'),
			dataMashupXml,
			'utf8'
		);
		log(`Debug: Saved original DataMashup XML to ${debugDir}/original_datamashup.xml`, 'debug');
		
		// Use excel-datamashup to correctly update the DataMashup binary content
		try {
			log('Attempting to parse existing DataMashup with excel-datamashup...');
			// Parse the existing DataMashup to get structure
			const parseResult = await excelDataMashup.ParseXml(dataMashupXml);
			
			if (typeof parseResult === 'string') {
				throw new Error(`Failed to parse existing DataMashup: ${parseResult}`);
			}
			
			log('DataMashup parsed successfully, updating formula...');
			// Use setFormula to update the M code (this also calls resetPermissions)
			parseResult.setFormula(cleanMCode);
			
			log('Formula updated, generating new DataMashup content...');
			// Use save to get the updated base64 binary content
			const newBase64Content = await parseResult.save();
			
			log(`excel-datamashup save() returned type: ${typeof newBase64Content}, length: ${String(newBase64Content).length}`);
			
			if (typeof newBase64Content === 'string' && newBase64Content.length > 0) {
				log('✅ excel-datamashup approach succeeded, updating Excel file...');
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
				log(`Successfully synced Power Query to Excel: ${path.basename(excelFile)}`);
				
				// Open Excel after sync if enabled
				const config = getConfig();
				if (config.get<boolean>('sync.openExcelAfterWrite', false)) {
					try {
						await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(excelFile));
						log(`Opened Excel file after sync: ${path.basename(excelFile)}`);
					} catch (openError) {
						log(`Failed to open Excel file after sync: ${openError}`, "error");
					}
				}
				return;
				
			} else {
				throw new Error(`excel-datamashup save() returned invalid content - Type: ${typeof newBase64Content}, Length: ${String(newBase64Content).length}`);
			}
			
		} catch (dataMashupError) {
			log(`❌ excel-datamashup approach failed: ${dataMashupError}`, "error");
			throw new Error(`DataMashup sync failed: ${dataMashupError}. The DataMashup format may have changed or be unsupported.`);
		}
		
	} catch (error) {
		const errorMsg = `Failed to sync to Excel: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
		console.error('Sync error:', error);
		
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
					log(`Restored from backup: ${backupPath}`);
				}
			}
		}
	}
}

async function watchFile(uri?: vscode.Uri): Promise<void> {
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			vscode.window.showErrorMessage('Please select or open a .m file to watch.');
			return;
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
				log('Test environment: Missing Excel file, proceeding with watch anyway', 'watchFile');
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
	log(`Setting up file watcher for: ${mFile}`, 'watchFile');
	log(`Remote environment: ${vscode.env.remoteName}`, 'watchFile');
	log(`Is dev container: ${vscode.env.remoteName === 'dev-container'}`, 'watchFile');
	
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
	
	log(`Chokidar watcher created for ${path.basename(mFile)}, polling: ${isDevContainer}`, 'watchFile');
	
	// Add comprehensive event logging
	watcher.on('change', async () => {
		try {
			log(`🔥 CHOKIDAR: File change detected: ${path.basename(mFile)}`, 'watchFile');
			vscode.window.showInformationMessage(`📝 File changed, syncing: ${path.basename(mFile)}`);
			log(`File changed, triggering debounced sync: ${path.basename(mFile)}`, 'watchFile');
			debouncedSyncToExcel(mFile);
		} catch (error) {
			const errorMsg = `Auto-sync failed: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, "watchFile");
		}
	});
	
	watcher.on('add', (path) => {
		log(`🆕 CHOKIDAR: File added: ${path}`, 'watchFile');
		// DON'T trigger sync on file creation - only on user changes
	});
	
	watcher.on('unlink', (path) => {
		log(`🗑️ CHOKIDAR: File deleted: ${path}`, 'watchFile');
	});
	
	watcher.on('error', (error) => {
		log(`❌ CHOKIDAR: Watcher error: ${error}`, 'watchFile');
	});
	
	watcher.on('ready', () => {
		log(`✅ CHOKIDAR: Watcher ready for ${path.basename(mFile)}`, 'watchFile');
	});

	// BACKUP WATCHER: Only add VS Code FileSystemWatcher in dev containers as backup
	let vscodeWatcher: vscode.FileSystemWatcher | undefined;
	let documentWatcher: vscode.Disposable | undefined;
	
	if (isDevContainer) {
		log(`Adding backup watchers for dev container environment`, 'watchFile');
		
		vscodeWatcher = vscode.workspace.createFileSystemWatcher(mFile);
		vscodeWatcher.onDidChange(async () => {
			try {
				log(`🔥 VSCODE: File change detected: ${path.basename(mFile)}`, 'watchFile');
				vscode.window.showInformationMessage(`📝 File changed (VSCode watcher), syncing: ${path.basename(mFile)}`);
				debouncedSyncToExcel(mFile);
			} catch (error) {
				log(`VS Code watcher sync failed: ${error}`, 'watchFile');
			}
		});
		
		vscodeWatcher.onDidCreate(() => {
			log(`🆕 VSCODE: File created: ${path.basename(mFile)}`, 'watchFile');
		});
		
		vscodeWatcher.onDidDelete(() => {
			log(`🗑️ VSCODE: File deleted: ${path.basename(mFile)}`, 'watchFile');
		});

		log(`VS Code FileSystemWatcher created for ${path.basename(mFile)}`, 'watchFile');

		// EXPERIMENTAL: Document save events as additional trigger (dev container only)
		documentWatcher = vscode.workspace.onDidSaveTextDocument(async (document) => {
			if (document.fileName === mFile) {
				try {
					log(`💾 DOCUMENT: Save event detected: ${path.basename(mFile)}`, 'watchFile');
					vscode.window.showInformationMessage(`📝 File saved (document event), syncing: ${path.basename(mFile)}`);
					debouncedSyncToExcel(mFile);
				} catch (error) {
					log(`Document save event sync failed: ${error}`, 'watchFile');
				}
			}
		});
		
		log(`VS Code document save watcher created for ${path.basename(mFile)}`, 'watchFile');
	} else {
		log(`Windows environment detected - using Chokidar only to avoid cascade events`, 'watchFile');
	}		// Store watchers for cleanup (handle optional backup watchers)
		const watcherSet = { 
			chokidar: watcher, 
			vscode: vscodeWatcher || null,
			document: documentWatcher || null
		};
		fileWatchers.set(mFile, watcherSet);
		
		const excelFileName = excelFile ? path.basename(excelFile) : 'Excel file (when found)';
		vscode.window.showInformationMessage(`👀 Now watching: ${path.basename(mFile)} → ${excelFileName}`);
		log(`Started watching: ${path.basename(mFile)}`);
		updateStatusBar();
		
		// Ensure the Promise resolves after watchers are set up
		return Promise.resolve();
		
	} catch (error) {
		const errorMsg = `Failed to watch file: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
		console.error('Watch error:', error);
	}
}

async function toggleWatch(uri?: vscode.Uri): Promise<void> {
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			vscode.window.showErrorMessage('Please select or open a .m file to toggle watch.');
			return;
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
		log(errorMsg, "error");
		console.error('Toggle watch error:', error);
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
		log(`Stopped watching: ${path.basename(mFile)}`);
		updateStatusBar();
	} else {
		vscode.window.showInformationMessage(`File was not being watched: ${path.basename(mFile)}`);
	}
}

async function syncAndDelete(uri?: vscode.Uri): Promise<void> {
	try {
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (!mFile || !mFile.endsWith('.m')) {
			vscode.window.showErrorMessage('Please select or open a .m file to sync and delete.');
			return;
		}

		const config = getConfig();
		let confirmation: string | undefined = 'Yes, Sync & Delete';
		
		// Ask for confirmation if setting is enabled
		if (config.get<boolean>('syncDeleteAlwaysConfirm', true)) {
			confirmation = await vscode.window.showWarningMessage(
				`Sync ${path.basename(mFile)} to Excel and then delete the .m file?`,
				{ modal: true },
				'Yes, Sync & Delete', 'Cancel'
			);
		}
		
		if (confirmation === 'Yes, Sync & Delete') {
			// First try to sync
			try {
				await syncToExcel(uri);
				
				// Stop watching if enabled and if being watched
				const watchers = fileWatchers.get(mFile);
				if (watchers) {
					if (config.get<boolean>('syncDeleteTurnsWatchOff', true)) {
						await watchers.chokidar.close();
						watchers.vscode?.dispose();
						watchers.document?.dispose();
						fileWatchers.delete(mFile);
						log(`Stopped watching due to sync & delete: ${path.basename(mFile)}`);
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
				log(`Successfully synced and deleted: ${path.basename(mFile)}`);
				
			} catch (syncError) {
				const errorMsg = `Sync failed, file not deleted: ${syncError}`;
				vscode.window.showErrorMessage(errorMsg);
				log(errorMsg, "error");
			}
		}
	} catch (error) {
		const errorMsg = `Sync and delete failed: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
		console.error('Sync and delete error:', error);
	}
}

async function rawExtraction(uri?: vscode.Uri): Promise<void> {
	try {
		// Dump extension settings for debugging (debug level only)
		const logLevel = getEffectiveLogLevel();
		if (logLevel === 'debug') {
			dumpAllExtensionSettings();
		}
		
		const excelFile = uri?.fsPath || await selectExcelFile();
		if (!excelFile) {
			return;
		}

		log(`Starting enhanced raw extraction for: ${path.basename(excelFile)}`);
		
		// Create debug output directory (delete if exists)
		const baseName = path.basename(excelFile, path.extname(excelFile));
		const outputDir = path.join(path.dirname(excelFile), `${baseName}_debug_extraction`);
		
		// Clean up existing debug directory
		if (fs.existsSync(outputDir)) {
			log(`Cleaning up existing debug directory: ${outputDir}`);
			fs.rmSync(outputDir, { recursive: true, force: true });
		}
		fs.mkdirSync(outputDir);
		log(`Created fresh debug directory: ${outputDir}`);

		// Get file stats
		const fileStats = fs.statSync(excelFile);
		const fileSizeMB = (fileStats.size / (1024 * 1024)).toFixed(2);
		log(`File size: ${fileSizeMB} MB`);

		// Use JSZip to extract and examine the Excel file structure
		try {
			const JSZip = (await import('jszip')).default;
			log('Reading Excel file buffer...');
			const buffer = fs.readFileSync(excelFile);
			
			log('Loading ZIP structure...');
			const startTime = Date.now();
			const zip = await JSZip.loadAsync(buffer);
			const loadTime = Date.now() - startTime;
			log(`ZIP loaded in ${loadTime}ms`);
			
			// List all files
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			log(`Found ${allFiles.length} files in ZIP structure`);
			
			// Categorize files
			const customXmlFiles = allFiles.filter(f => f.startsWith('customXml/'));
			const xlFiles = allFiles.filter(f => f.startsWith('xl/'));
			const queryFiles = allFiles.filter(f => f.includes('quer') || f.includes('Query'));
			const connectionFiles = allFiles.filter(f => f.includes('connection'));
			
			log(`Files breakdown: ${customXmlFiles.length} customXml, ${xlFiles.length} xl/, ${queryFiles.length} query-related, ${connectionFiles.length} connection-related`);

			// Enhanced DataMashup detection - scan ALL XML files
			const dataMashupResults: Array<{file: string, hasDataMashup: boolean, size: number, error?: string}> = [];
			const xmlFiles = allFiles.filter(f => f.toLowerCase().endsWith('.xml'));
			
			log(`Scanning ${xmlFiles.length} XML files for DataMashup content...`);
			
			for (const fileName of xmlFiles) {
				try {
					const file = zip.file(fileName);
					if (file) {
						const content = await file.async('text');
						const hasDataMashup = content.includes('<DataMashup') || content.includes('DataMashup');
						const size = content.length;
						
						dataMashupResults.push({
							file: fileName,
							hasDataMashup,
							size
						});
						
						if (hasDataMashup) {
							log(`✅ DataMashup found in: ${fileName} (${(size / 1024).toFixed(1)} KB)`);
							
							// Extract the DataMashup content to a separate file
							const safeName = fileName.replace(/[\/\\]/g, '_');
							const dataMashupPath = path.join(outputDir, `DATAMASHUP_${safeName}`);
							fs.writeFileSync(dataMashupPath, content, 'utf8');
							log(`📁 DataMashup extracted to: ${path.basename(dataMashupPath)}`);
						}
						
						// Extract customXml files regardless for inspection
						if (fileName.startsWith('customXml/')) {
							const safeName = fileName.replace(/[\/\\]/g, '_');
							fs.writeFileSync(
								path.join(outputDir, `${safeName}.txt`),
								content,
								'utf8'
							);
						}
					}
				} catch (error) {
					log(`❌ Error scanning ${fileName}: ${error}`);
					dataMashupResults.push({
						file: fileName,
						hasDataMashup: false,
						size: 0,
						error: String(error)
					});
				}
			}
			
			// Count DataMashup findings
			const dataMashupFiles = dataMashupResults.filter(r => r.hasDataMashup);
			const totalDataMashupSize = dataMashupFiles.reduce((sum, r) => sum + r.size, 0);
			
			log(`DataMashup scan complete: Found ${dataMashupFiles.length} files containing DataMashup (${(totalDataMashupSize / 1024).toFixed(1)} KB total)`);

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
					totalXmlFilesScanned: xmlFiles.length,
					dataMashupFilesFound: dataMashupFiles.length,
					totalDataMashupSize: `${(totalDataMashupSize / 1024).toFixed(1)} KB`,
					results: dataMashupResults
				},
				potentialPowerQueryLocations: customXmlFiles.concat([
					'xl/queryTables/queryTable1.xml',
					'xl/connections.xml'
				]).filter(loc => allFiles.includes(loc)),
				recommendations: dataMashupFiles.length === 0 ? 
					['No DataMashup content found - file may not contain Power Query M code', 'Check if Excel file actually has Power Query connections'] :
					[`Found DataMashup in: ${dataMashupFiles.map(f => f.file).join(', ')}`, 'Use extracted DataMashup files for further analysis']
			};

			const reportPath = path.join(outputDir, 'EXTRACTION_REPORT.json');
			fs.writeFileSync(reportPath, JSON.stringify(debugInfo, null, 2), 'utf8');
			log(`📊 Comprehensive report saved: ${path.basename(reportPath)}`);

			// Show results
			const message = dataMashupFiles.length > 0 ?
				`✅ Enhanced extraction completed!\n🔍 Found ${dataMashupFiles.length} DataMashup source(s) in ${path.basename(excelFile)}\n📁 Results in: ${path.basename(outputDir)}` :
				`⚠️ Enhanced extraction completed!\n❌ No DataMashup content found in ${path.basename(excelFile)}\n📁 Debug files in: ${path.basename(outputDir)}`;
			
			vscode.window.showInformationMessage(message);
			log(message.replace(/\n/g, ' | '));
			
		} catch (error) {
			log(`❌ ZIP extraction/analysis failed: ${error}`, "error");
			
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
		log(errorMsg, "error");
		console.error('Raw extraction error:', error);
	}
}

// New function to dump all extension settings for debugging
function dumpAllExtensionSettings(): void {
	try {
		log('=== EXTENSION SETTINGS DUMP ===');
		
		const extensionId = 'excel-power-query-editor';
		
		// Get all configuration scopes
		const userConfig = vscode.workspace.getConfiguration(extensionId, null);
		const workspaceConfig = vscode.workspace.getConfiguration(extensionId, vscode.workspace.workspaceFolders?.[0]?.uri);
		
		// Define all known extension settings
		const knownSettings = [
			'watchAlways',
			'watchOffOnDelete', 
			'syncDeleteAlwaysConfirm',
			'verboseMode',
			'autoBackupBeforeSync',
			'backupLocation',
			'customBackupPath',
			'backup.maxFiles',
			'autoCleanupBackups',
			'syncTimeout',
			'debugMode',
			'showStatusBarInfo',
			'sync.openExcelAfterWrite',
			'sync.debounceMs',
			'watch.checkExcelWriteable'
		];
		
		log('USER SETTINGS (Global):');
		for (const setting of knownSettings) {
			const value = userConfig.get(setting);
			const hasValue = userConfig.has(setting);
			log(`  ${setting}: ${hasValue ? JSON.stringify(value) : '<not set>'}`);
		}
		
		log('WORKSPACE SETTINGS:');
		for (const setting of knownSettings) {
			const value = workspaceConfig.get(setting);
			const hasValue = workspaceConfig.has(setting);
			log(`  ${setting}: ${hasValue ? JSON.stringify(value) : '<not set>'}`);
		}
		
		// Check environment info
		log('ENVIRONMENT INFO:');
		log(`  Remote Name: ${vscode.env.remoteName || '<not remote>'}`);
		log(`  VS Code Version: ${vscode.version}`);
		log(`  Workspace Folders: ${vscode.workspace.workspaceFolders?.length || 0}`);
		
		// Check if we're in a dev container
		const isDevContainer = vscode.env.remoteName?.includes('dev-container');
		log(`  Is Dev Container: ${isDevContainer}`);
		
		log('=== END SETTINGS DUMP ===');
		
	} catch (error) {
		log(`Failed to dump settings: ${error}`, "error");
	}
}

async function selectExcelFile(): Promise<string | undefined> {
	// In test environment, return a test fixture instead of showing dialog
	if (isTestEnvironment()) {
		const testFixtures = ['simple.xlsx', 'complex.xlsm', 'binary.xlsb'];
		for (const fixture of testFixtures) {
			const fixturePath = getTestFixturePath(fixture);
			if (fs.existsSync(fixturePath)) {
				log(`Test environment: Using fixture ${fixture}`, 'selectExcelFile');
				return fixturePath;
			}
		}
		log('Test environment: No fixtures found, returning undefined', 'selectExcelFile');
		return undefined;
	}
	
	// Normal user interaction for production
	const result = await vscode.window.showOpenDialog({
		canSelectFiles: true,
		canSelectFolders: false,
		canSelectMany: false,
		filters: {
			'Excel Files': ['xlsx', 'xlsm', 'xlsb']
		}
	});
	
	return result?.[0]?.fsPath;
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
		const excelFile = uri?.fsPath || await selectExcelFile();
		if (!excelFile) {
			return;
		}

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
			const originalAutoCleanup = config.get<boolean>('autoCleanupBackups', true);
			if (config.update) {
				await config.update('autoCleanupBackups', true, vscode.ConfigurationTarget.Global);
			}
			
			try {
				cleanupOldBackups(excelFile);
				vscode.window.showInformationMessage(`✅ Backup cleanup completed for ${path.basename(excelFile)}`);
			} finally {
				// Restore original setting
				if (config.update) {
					await config.update('autoCleanupBackups', originalAutoCleanup, vscode.ConfigurationTarget.Global);
				}
			}
		}
		
	} catch (error) {
		const errorMsg = `Failed to cleanup backups: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
		console.error('Backup cleanup error:', error);
	}
}

// Apply recommended default settings for v0.5.0
async function applyRecommendedDefaults(): Promise<void> {
	try {
		const config = vscode.workspace.getConfiguration('excel-power-query-editor');
		
		// Recommended settings for v0.5.0 (using new logLevel instead of legacy boolean flags)
		const recommendedSettings = {
			'watchAlways': false,
			'watchOffOnDelete': true,
			'syncDeleteAlwaysConfirm': true,
			'logLevel': 'info', // New setting replaces verboseMode and debugMode
			'autoBackupBeforeSync': true,
			'backupLocation': 'sameFolder',
			'backup.maxFiles': 5,
			'autoCleanupBackups': true,
			'syncTimeout': 30000,
			'showStatusBarInfo': true,
			'sync.openExcelAfterWrite': false,
			'sync.debounceMs': 500,
			'watch.checkExcelWriteable': true
		};
		
		let updatedCount = 0;
		const changedSettings: string[] = [];
		
		// Detect if running in dev container for workspace vs global scope
		const isDevContainer = vscode.env.remoteName?.includes("dev-container") || 
							   vscode.env.remoteName?.includes("container") ||
							   process.env.REMOTE_CONTAINERS === "true";
		const configTarget = isDevContainer ? vscode.ConfigurationTarget.Workspace : vscode.ConfigurationTarget.Global;
		
		for (const [setting, value] of Object.entries(recommendedSettings)) {
			const currentValue = config.get(setting);
			if (currentValue !== value) {
				await config.update(setting, value, configTarget);
				changedSettings.push(`${setting}: ${currentValue} → ${value}`);
				updatedCount++;
			}
		}
		
		// Clean up legacy settings if present (optional migration)
		const legacySettings = ['verboseMode', 'debugMode'];
		for (const legacySetting of legacySettings) {
			const legacyValue = config.get(legacySetting);
			if (legacyValue !== undefined) {
				try {
					await config.update(legacySetting, undefined, configTarget);
					changedSettings.push(`${legacySetting}: ${legacyValue} → (removed - use logLevel instead)`);
					updatedCount++;
				} catch (cleanupError) {
					log(`Could not remove legacy setting ${legacySetting}: ${cleanupError}`, 'warn');
				}
			}
		}
		
		if (updatedCount > 0) {
			const scope = isDevContainer ? 'workspace' : 'global';
			vscode.window.showInformationMessage(
				`✅ Applied recommended defaults for v0.5.0 (${updatedCount} settings updated in ${scope} scope)`
			);
			log(`Applied recommended defaults (${scope} scope) - Updated settings:\n${changedSettings.join('\n')}`);
		} else {
			vscode.window.showInformationMessage(
				'All settings already match recommended defaults for v0.5.0'
			);
			log('All settings already match recommended defaults');
		}
		
	} catch (error) {
		const errorMsg = `Failed to apply recommended defaults: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, "error");
	}
}

// Debounced sync helper to prevent multiple syncs in rapid succession
function debouncedSyncToExcel(mFile: string): void {
	// Check if this file was recently extracted - if so, skip auto-sync
	if (recentExtractions.has(mFile)) {
		log(`⏭️ Skipping auto-sync for recently extracted file: ${path.basename(mFile)}`, 'debouncedSyncToExcel');
		return;
	}
	
	const config = getConfig();
	const debounceMs = config.get<number>('sync.debounceMs', 500);
	
	// If debounce is 0 or minimal (100ms), execute immediately for debugging
	if (debounceMs === 0 || (debounceMs && debounceMs <= 100)) {
		log(`🚀 IMMEDIATE SYNC (debounce disabled: ${debounceMs}ms) for ${path.basename(mFile)}`, 'debouncedSyncToExcel');
		syncToExcel(vscode.Uri.file(mFile)).catch(error => {
			log(`Immediate sync failed for ${path.basename(mFile)}: ${error}`, "error");
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
			log(`Debounced sync executing for ${path.basename(mFile)}`);
			await syncToExcel(vscode.Uri.file(mFile));
			debounceTimers.delete(mFile);
		} catch (error) {
			log(`Debounced sync failed for ${path.basename(mFile)}: ${error}`, "error");
			debounceTimers.delete(mFile);
		}
	}, debounceMs);
	
	debounceTimers.set(mFile, timer);
	log(`Sync debounced for ${path.basename(mFile)} (${debounceMs}ms)`);
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
		log(`Excel file appears to be locked: ${error.message}`, "error");
		return false;
	}
}

// This method is called when your extension is deactivated
export function deactivate() {
	// Close all file watchers
	for (const [, watchers] of fileWatchers) {
		watchers.chokidar.close();
		watchers.vscode?.dispose();
		watchers.document?.dispose();
	}
	fileWatchers.clear();
}

// Parse structured metadata from .m file header
