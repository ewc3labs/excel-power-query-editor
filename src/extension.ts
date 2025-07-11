// The module 'vscode' contains the VS Code extensibility API
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { watch, FSWatcher } from 'chokidar';
import { getConfig } from './configHelper';

// File watchers storage
const fileWatchers = new Map<string, FSWatcher>();

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
					log(`Failed to delete backup ${backup.filename}: ${deleteError}`, true);
				}
			}
			
			if (deletedCount > 0) {
				log(`Cleaned up ${deletedCount} old backup files (keeping ${maxBackups} most recent)`);
			}
		}
		
	} catch (error) {
		log(`Backup cleanup failed: ${error}`, true);
	}
}

// Verbose logging helper
function log(message: string, isError: boolean = false) {
	const config = getConfig();
	const timestamp = new Date().toISOString();
	const logMessage = `[${timestamp}] ${message}`;
	
	console.log(logMessage);
	
	if (config.get<boolean>('verboseMode', false)) {
		if (!outputChannel) {
			outputChannel = vscode.window.createOutputChannel('Excel Power Query Editor');
		}
		outputChannel.appendLine(logMessage);
		if (isError) {
			outputChannel.show();
		}
	}
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
					log(`Failed to auto-watch ${path.basename(mFile)}: ${error}`, true);
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
		log(`Auto-watch initialization failed: ${error}`, true);
		vscode.window.showErrorMessage(`Auto-watch initialization failed: ${error}`);
	}
}

// This method is called when your extension is activated
export async function activate(context: vscode.ExtensionContext) {
	console.log('Excel Power Query Editor extension is now active!');

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

	// Initialize output channel and status bar
	outputChannel = vscode.window.createOutputChannel('Excel Power Query Editor');
	updateStatusBar();
	
	log('Excel Power Query Editor extension activated');
	
	// Auto-watch existing .m files if setting is enabled
	await initializeAutoWatch();
}

async function extractFromExcel(uri?: vscode.Uri): Promise<void> {
	try {
		const excelFile = uri?.fsPath || await selectExcelFile();
		if (!excelFile) {
			return;
		}

		vscode.window.showInformationMessage(`Extracting Power Query from: ${path.basename(excelFile)}`);
		
		// Try to use excel-datamashup for extraction
		try {
			// First, we need to extract customXml/item1.xml from the Excel file
			const JSZip = (await import('jszip')).default;
			
			// Use require for excel-datamashup to avoid ES module issues
			const excelDataMashup = require('excel-datamashup');
			
			const buffer = fs.readFileSync(excelFile);
			const zip = await JSZip.loadAsync(buffer);
			
			// Debug: List all files in the Excel zip
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			console.log('Files in Excel archive:', allFiles);
			
			// Look for Power Query in multiple possible locations
			const powerQueryLocations = [
				'customXml/item1.xml',
				'customXml/item2.xml', 
				'customXml/item3.xml',
				'xl/queryTables/queryTable1.xml',
				'xl/connections.xml'
			];
			
			let xmlContent: string | null = null;
			let foundLocation = '';
			let queryType = '';
			
			for (const location of powerQueryLocations) {
				const xmlFile = zip.file(location);
				if (xmlFile) {
					try {
						// Read as binary first, then decode properly
						const binaryData = await xmlFile.async('nodebuffer');
						let content: string;
						
						// Check for UTF-16 LE BOM (FF FE)
						if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
							console.log(`Detected UTF-16 LE BOM in ${location}`);
							// Decode UTF-16 LE (skip the 2-byte BOM)
							content = binaryData.subarray(2).toString('utf16le');
						} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
							console.log(`Detected UTF-8 BOM in ${location}`);
							// Decode UTF-8 (skip the 3-byte BOM)
							content = binaryData.subarray(3).toString('utf8');
						} else {
							// Try UTF-8 first (most common)
							content = binaryData.toString('utf8');
						}
						
						console.log(`Content preview from ${location} (first 200 chars):`, content.substring(0, 200));
						
						// Check for DataMashup format (what excel-datamashup expects)
						if (content.includes('DataMashup')) {
							xmlContent = content;
							foundLocation = location;
							queryType = 'DataMashup';
							console.log(`Found DataMashup Power Query in: ${location}`);
							break;
						}
						// Check for query table format (newer Excel)
						else if (content.includes('queryTable') && location.includes('queryTables')) {
							xmlContent = content;
							foundLocation = location;
							queryType = 'QueryTable';
							console.log(`Found QueryTable Power Query in: ${location}`);
							break;
						}
						// Check for connections format
						else if (content.includes('connection') && (content.includes('Query') || content.includes('PowerQuery'))) {
							xmlContent = content;
							foundLocation = location;
							queryType = 'Connection';
							console.log(`Found Connection Power Query in: ${location}`);
							break;
						}
					} catch (e) {
						console.log(`Could not read ${location}:`, e);
					}
				}
			}
			
			if (!xmlContent) {
				// No Power Query found, let's check what customXml files exist
				const customXmlFiles = allFiles.filter(f => f.startsWith('customXml/'));
				const xlFiles = allFiles.filter(f => f.startsWith('xl/') && f.includes('quer'));
				
				vscode.window.showWarningMessage(
					`No Power Query found. Available files:\n` +
					`CustomXml: ${customXmlFiles.join(', ') || 'none'}\n` +
					`Query files: ${xlFiles.join(', ') || 'none'}\n` +
					`Total files: ${allFiles.length}`
				);
				return;
			}
			
			console.log(`Attempting to parse Power Query from: ${foundLocation} (type: ${queryType})`);
			
			if (queryType === 'DataMashup') {
				// Use excel-datamashup for DataMashup format
				const parseResult = await excelDataMashup.ParseXml(xmlContent);
				
				if (typeof parseResult === 'string') {
					vscode.window.showErrorMessage(`Power Query parsing failed: ${parseResult}\nLocation: ${foundLocation}\nXML preview: ${xmlContent.substring(0, 200)}...`);
					return;
				}
				
				// Extract the formula
				const formula = parseResult.getFormula();
				if (!formula) {
					vscode.window.showWarningMessage(`No Power Query formula found in ${foundLocation}. ParseResult keys: ${Object.keys(parseResult).join(', ')}`);
					return;
				}
				
				// Create output file with the actual formula
				const baseName = path.basename(excelFile);
				const outputPath = path.join(path.dirname(excelFile), `${baseName}_PowerQuery.m`);
				
				const content = `// Power Query extracted from: ${path.basename(excelFile)}
// Location: ${foundLocation} (DataMashup format)
// Extracted on: ${new Date().toISOString()}

${formula}`;

				fs.writeFileSync(outputPath, content, 'utf8');
				
				// Open the created file
				const document = await vscode.workspace.openTextDocument(outputPath);
				await vscode.window.showTextDocument(document);
				
				vscode.window.showInformationMessage(`Power Query extracted to: ${path.basename(outputPath)}`);
				log(`Successfully extracted Power Query from ${path.basename(excelFile)} to ${path.basename(outputPath)}`);
				
				// Auto-watch if enabled
				const config = getConfig();
				if (config.get<boolean>('watchAlways', false)) {
					await watchFile(vscode.Uri.file(outputPath));
					log(`Auto-watch enabled for ${path.basename(outputPath)}`);
				}
				
			} else {
				// Handle QueryTable or Connection format (extract what we can)
				const baseName = path.basename(excelFile);
				const outputPath = path.join(path.dirname(excelFile), `${baseName}_PowerQuery.m`);
				
				let extractedContent = '';
				
				if (queryType === 'QueryTable') {
					// Try to extract useful information from query table XML
					const connectionMatch = xmlContent.match(/<connectionId>(\d+)<\/connectionId>/);
					const nameMatch = xmlContent.match(/name="([^"]+)"/);
					
					extractedContent = `// Power Query extracted from: ${path.basename(excelFile)}
// Location: ${foundLocation} (QueryTable format)
// Extracted on: ${new Date().toISOString()}
//
// Note: This is a QueryTable format, not full Power Query M code.
// Connection ID: ${connectionMatch ? connectionMatch[1] : 'unknown'}
// Table Name: ${nameMatch ? nameMatch[1] : 'unknown'}
//
// TODO: Full M code extraction not yet supported for this format.
// Raw XML content below for reference:

/*
${xmlContent}
*/

let
    // Placeholder - actual query needs to be reconstructed
    Source = Excel.CurrentWorkbook(){[Name="${nameMatch ? nameMatch[1] : 'Table1'}"]}[Content],
    Result = Source
in
    Result`;
				} else {
					extractedContent = `// Power Query extracted from: ${path.basename(excelFile)}
// Location: ${foundLocation} (${queryType} format)
// Extracted on: ${new Date().toISOString()}
//
// Note: This format is not fully supported yet.
// Raw XML content below for reference:

/*
${xmlContent}
*/

let
    // Placeholder - actual query needs to be reconstructed
    Source = "Power Query data found but format not yet supported",
    Result = Source
in
    Result`;
				}
				
				fs.writeFileSync(outputPath, extractedContent, 'utf8');
				
				// Open the created file
				const document = await vscode.workspace.openTextDocument(outputPath);
				await vscode.window.showTextDocument(document);
				
				vscode.window.showInformationMessage(`Power Query partially extracted to: ${path.basename(outputPath)} (${queryType} format - limited support)`);
				log(`Partially extracted Power Query from ${path.basename(excelFile)} to ${path.basename(outputPath)} (${queryType} format)`);
				
				// Auto-watch if enabled
				const config = getConfig();
				if (config.get<boolean>('watchAlways', false)) {
					await watchFile(vscode.Uri.file(outputPath));
					log(`Auto-watch enabled for ${path.basename(outputPath)}`);
				}
			}

			// ...existing code...
		} catch (moduleError) {
			// Fallback: create a placeholder file
			vscode.window.showWarningMessage(`Excel parsing failed: ${moduleError}. Creating placeholder file for testing.`);
			
			const baseName = path.basename(excelFile); // Keep full filename including extension
			const outputPath = path.join(path.dirname(excelFile), `${baseName}_PowerQuery.m`);
			
			const placeholderContent = `// Power Query extraction from: ${path.basename(excelFile)}
// 
// This is a placeholder file - actual extraction failed.
// Error: ${moduleError}
//
// File: ${excelFile}
// Extracted on: ${new Date().toISOString()}
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
		log(errorMsg, true);
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

		// Find corresponding Excel file
		let excelFile = await findExcelFile(mFile);
		if (!excelFile) {
			vscode.window.showErrorMessage('Could not find corresponding Excel file. Please select one.');
			const selected = await selectExcelFile();
			if (!selected) {
				return;
			}
			excelFile = selected;
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
		
		// Extract just the M code (remove our comment headers)
		const mCodeMatch = mContent.match(/(?:\/\/.*\n)*\n*([\s\S]+)/);
		const cleanMCode = mCodeMatch ? mCodeMatch[1].trim() : mContent.trim();
		
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
		
		// Find the DataMashup XML file
		let dataMashupFile = zip.file('customXml/item1.xml');
		if (!dataMashupFile) {
			vscode.window.showErrorMessage('No DataMashup found in Excel file. This file may not contain Power Query.');
			return;
		}
		
		// Read and decode the DataMashup XML
		const binaryData = await dataMashupFile.async('nodebuffer');
		let dataMashupXml: string;
		
		// Handle UTF-16 LE BOM like in extraction
		if (binaryData.length >= 2 && binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
			console.log('Detected UTF-16 LE BOM in DataMashup');
			dataMashupXml = binaryData.subarray(2).toString('utf16le');
		} else if (binaryData.length >= 3 && binaryData[0] === 0xEF && binaryData[1] === 0xBB && binaryData[2] === 0xBF) {
			console.log('Detected UTF-8 BOM in DataMashup');
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
		console.log(`Debug: Saved original DataMashup XML to ${debugDir}/original_datamashup.xml`);
		
		// Use excel-datamashup to correctly update the DataMashup binary content
		try {
			// Parse the existing DataMashup to get structure
			const parseResult = await excelDataMashup.ParseXml(dataMashupXml);
			
			if (typeof parseResult === 'string') {
				throw new Error(`Failed to parse existing DataMashup: ${parseResult}`);
			}
			
			// Use setFormula to update the M code (this also calls resetPermissions)
			parseResult.setFormula(cleanMCode);
			
			// Use save to get the updated base64 binary content
			const newBase64Content = await parseResult.save();
			
			// DEBUG: Save the result from excel-datamashup save()
			fs.writeFileSync(
				path.join(debugDir, 'excel_datamashup_save_result.txt'),
				`Type: ${typeof newBase64Content}\nContent: ${String(newBase64Content).substring(0, 1000)}...`,
				'utf8'
			);
			console.log(`Debug: excel-datamashup save() returned type: ${typeof newBase64Content}`);
			
			if (typeof newBase64Content === 'string' && newBase64Content.length > 0) {
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
				
				// Update the ZIP with new DataMashup
				zip.file('customXml/item1.xml', newBinaryData);
				
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
						log(`Failed to open Excel file after sync: ${openError}`, true);
					}
				}
				return;
				
			} else {
				throw new Error('excel-datamashup save() failed or returned empty content');
			}
			
		} catch (dataMashupError) {
			console.log('excel-datamashup approach failed, trying manual XML modification:', dataMashupError);
			
			// Fallback: Manual XML modification using xml2js
			try {
				const parser = new xml2js.Parser();
				const builder = new xml2js.Builder({ 
					renderOpts: { pretty: false },
					xmldec: { version: '1.0', encoding: 'utf-16' }
				});
				
				const parsedXml = await parser.parseStringPromise(dataMashupXml);
				
				// DEBUG: Save the parsed XML structure
				fs.writeFileSync(
					path.join(debugDir, 'parsed_xml_structure.json'),
					JSON.stringify(parsedXml, null, 2),
					'utf8'
				);
				console.log(`Debug: Saved parsed XML structure to ${debugDir}/parsed_xml_structure.json`);
				
				// Find and update the Formula section in the XML
				// This is a simplified approach - the actual structure may be more complex
				if (parsedXml.DataMashup && parsedXml.DataMashup.Formulas) {
					// Replace the entire Formulas section with our new M code
					// Note: This is a basic implementation and may need refinement
					parsedXml.DataMashup.Formulas = [{ _: cleanMCode }];
				} else {
					throw new Error(`Could not find Formulas section in DataMashup XML. Available sections: ${Object.keys(parsedXml.DataMashup || {}).join(', ')}`);
				}
				
				// Rebuild XML
				let newDataMashupXml = builder.buildObject(parsedXml);
				
				// Convert back to appropriate encoding
				let newBinaryData: Buffer;
				if (binaryData[0] === 0xFF && binaryData[1] === 0xFE) {
					const utf16Buffer = Buffer.from(newDataMashupXml, 'utf16le');
					const bomBuffer = Buffer.from([0xFF, 0xFE]);
					newBinaryData = Buffer.concat([bomBuffer, utf16Buffer]);
				} else {
					newBinaryData = Buffer.from(newDataMashupXml, 'utf8');
				}
				
				// Update the ZIP
				zip.file('customXml/item1.xml', newBinaryData);
				
				// Write the updated Excel file
				const updatedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
				fs.writeFileSync(excelFile, updatedBuffer);
				
				vscode.window.showInformationMessage(`✅ Successfully synced Power Query to Excel (manual method): ${path.basename(excelFile)}`);
				log(`Successfully synced Power Query to Excel (manual method): ${path.basename(excelFile)}`);
				
				// Open Excel after sync if enabled
				const config = getConfig();
				if (config.get<boolean>('sync.openExcelAfterWrite', false)) {
					try {
						await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(excelFile));
						log(`Opened Excel file after sync: ${path.basename(excelFile)}`);
					} catch (openError) {
						log(`Failed to open Excel file after sync: ${openError}`, true);
					}
				}
				
			} catch (manualError) {
				throw new Error(`Both excel-datamashup and manual XML approaches failed. DataMashup error: ${dataMashupError}. Manual error: ${manualError}`);
			}
		}
		
	} catch (error) {
		const errorMsg = `Failed to sync to Excel: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, true);
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
			const selection = await vscode.window.showWarningMessage(
				`Cannot find corresponding Excel file for ${path.basename(mFile)}. Watch anyway?`,
				'Yes, Watch Anyway', 'No'
			);
			if (selection !== 'Yes, Watch Anyway') {
				return;
			}
		}

		const watcher = watch(mFile, { 
			ignoreInitial: true,
			awaitWriteFinish: {
				stabilityThreshold: 300,
				pollInterval: 100
			}
		});
		
		watcher.on('change', async () => {
			try {
				vscode.window.showInformationMessage(`📝 File changed, syncing: ${path.basename(mFile)}`);
				log(`File changed, triggering debounced sync: ${path.basename(mFile)}`);
				debouncedSyncToExcel(mFile);
			} catch (error) {
				const errorMsg = `Auto-sync failed: ${error}`;
				vscode.window.showErrorMessage(errorMsg);
				log(errorMsg, true);
			}
		});
		
		watcher.on('unlink', () => {
			const config = getConfig();
			if (config.get<boolean>('watchOffOnDelete', true)) {
				fileWatchers.delete(mFile);
				log(`File deleted, stopped watching: ${path.basename(mFile)}`);
				updateStatusBar();
			}
		});
		
		watcher.on('error', (error) => {
			const errorMsg = `File watcher error: ${error}`;
			vscode.window.showErrorMessage(errorMsg);
			log(errorMsg, true);
			fileWatchers.delete(mFile);
			updateStatusBar();
		});

		fileWatchers.set(mFile, watcher);
		
		const excelFileName = excelFile ? path.basename(excelFile) : 'Excel file (when found)';
		vscode.window.showInformationMessage(`👀 Now watching: ${path.basename(mFile)} → ${excelFileName}`);
		log(`Started watching: ${path.basename(mFile)}`);
		updateStatusBar();
		
	} catch (error) {
		const errorMsg = `Failed to watch file: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, true);
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
		log(errorMsg, true);
		console.error('Toggle watch error:', error);
	}
}

async function stopWatching(uri?: vscode.Uri): Promise<void> {
	const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
	if (!mFile) {
		return;
	}

	const watcher = fileWatchers.get(mFile);
	if (watcher) {
		await watcher.close();
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
				const watcher = fileWatchers.get(mFile);
				if (watcher) {
					if (config.get<boolean>('syncDeleteTurnsWatchOff', true)) {
						await watcher.close();
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
				log(errorMsg, true);
			}
		}
	} catch (error) {
		const errorMsg = `Sync and delete failed: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, true);
		console.error('Sync and delete error:', error);
	}
}

async function rawExtraction(uri?: vscode.Uri): Promise<void> {
	try {
		const excelFile = uri?.fsPath || await selectExcelFile();
		if (!excelFile) {
			return;
		}

		// Create debug output directory
		const baseName = path.basename(excelFile, path.extname(excelFile));
		const outputDir = path.join(path.dirname(excelFile), `${baseName}_debug_extraction`);
		
		if (!fs.existsSync(outputDir)) {
			fs.mkdirSync(outputDir);
		}

		// Use JSZip to extract and examine the Excel file structure
		try {
			const JSZip = (await import('jszip')).default;
			const buffer = fs.readFileSync(excelFile);
			const zip = await JSZip.loadAsync(buffer);
			
			// List all files
			const allFiles = Object.keys(zip.files).filter(name => !zip.files[name].dir);
			
			// Look for potentially relevant files
			const customXmlFiles = allFiles.filter(f => f.startsWith('customXml/'));
			const xlFiles = allFiles.filter(f => f.startsWith('xl/'));
			const queryFiles = allFiles.filter(f => f.includes('quer') || f.includes('Query'));
			const connectionFiles = allFiles.filter(f => f.includes('connection'));
			
			// Extract customXml files for examination
			for (const fileName of customXmlFiles) {
				const file = zip.file(fileName);
				if (file) {
					const content = await file.async('text');
					const safeName = fileName.replace(/[\/\\]/g, '_');
					fs.writeFileSync(
						path.join(outputDir, `${safeName}.txt`),
						content,
						'utf8'
					);
				}
			}
			
			// Create a comprehensive debug report
			const debugInfo = {
				file: excelFile,
				extractedAt: new Date().toISOString(),
				totalFiles: allFiles.length,
				allFiles: allFiles,
				customXmlFiles: customXmlFiles,
				xlFiles: xlFiles,
				queryFiles: queryFiles,
				connectionFiles: connectionFiles,
				potentialPowerQueryLocations: [
					'customXml/item1.xml',
					'customXml/item2.xml', 
					'customXml/item3.xml',
					'xl/queryTables/queryTable1.xml',
					'xl/connections.xml'
				].filter(loc => allFiles.includes(loc))
			};

			fs.writeFileSync(
				path.join(outputDir, 'debug_info.json'),
				JSON.stringify(debugInfo, null, 2),
				'utf8'
			);

			vscode.window.showInformationMessage(`Debug extraction completed: ${path.basename(outputDir)}\nFound ${customXmlFiles.length} customXml files, ${queryFiles.length} query-related files`);
			
		} catch (error) {
			// Write error info
			const debugInfo = {
				file: excelFile,
				extractedAt: new Date().toISOString(),
				error: 'Failed to extract Excel file structure',
				errorDetails: String(error)
			};

			fs.writeFileSync(
				path.join(outputDir, 'debug_info.json'),
				JSON.stringify(debugInfo, null, 2),
				'utf8'
			);
		}
		
	} catch (error) {
		vscode.window.showErrorMessage(`Raw extraction failed: ${error}`);
		console.error('Raw extraction error:', error);
	}
}

async function selectExcelFile(): Promise<string | undefined> {
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
		log(errorMsg, true);
		console.error('Backup cleanup error:', error);
	}
}

// Apply recommended default settings for v0.5.0
async function applyRecommendedDefaults(): Promise<void> {
	try {
		const config = vscode.workspace.getConfiguration('excel-power-query-editor');
		
		// Recommended settings for v0.5.0
		const recommendedSettings = {
			'watchAlways': false,
			'watchOffOnDelete': true,
			'syncDeleteAlwaysConfirm': true,
			'verboseMode': false,
			'autoBackupBeforeSync': true,
			'backupLocation': 'sameFolder',
			'backup.maxFiles': 5,
			'autoCleanupBackups': true,
			'syncTimeout': 30000,
			'debugMode': false,
			'showStatusBarInfo': true,
			'sync.openExcelAfterWrite': false,
			'sync.debounceMs': 500,
			'watch.checkExcelWriteable': true
		};
		
		let updatedCount = 0;
		const changedSettings: string[] = [];
		
		for (const [setting, value] of Object.entries(recommendedSettings)) {
			const currentValue = config.get(setting);
			if (currentValue !== value) {
				await config.update(setting, value, vscode.ConfigurationTarget.Global);
				changedSettings.push(`${setting}: ${currentValue} → ${value}`);
				updatedCount++;
			}
		}
		
		if (updatedCount > 0) {
			vscode.window.showInformationMessage(
				`✅ Applied recommended defaults for v0.5.0 (${updatedCount} settings updated)`
			);
			log(`Applied recommended defaults - Updated settings:\n${changedSettings.join('\n')}`);
		} else {
			vscode.window.showInformationMessage(
				'All settings already match recommended defaults for v0.5.0'
			);
			log('All settings already match recommended defaults');
		}
		
	} catch (error) {
		const errorMsg = `Failed to apply recommended defaults: ${error}`;
		vscode.window.showErrorMessage(errorMsg);
		log(errorMsg, true);
	}
}

// Debounced sync helper to prevent multiple syncs in rapid succession
function debouncedSyncToExcel(mFile: string): void {
	const config = getConfig();
	const debounceMs = config.get<number>('sync.debounceMs', 500);
	
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
			log(`Debounced sync failed for ${path.basename(mFile)}: ${error}`, true);
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
		log(`Excel file appears to be locked: ${error.message}`, true);
		return false;
	}
}

// This method is called when your extension is deactivated
export function deactivate() {
	// Close all file watchers
	for (const [, watcher] of fileWatchers) {
		watcher.close();
	}
	fileWatchers.clear();
}
