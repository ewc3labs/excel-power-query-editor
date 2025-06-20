// The module 'vscode' contains the VS Code extensibility API
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { watch, FSWatcher } from 'chokidar';

// File watchers storage
const fileWatchers = new Map<string, FSWatcher>();

// Output channel for verbose logging
let outputChannel: vscode.OutputChannel;

// Status bar item for watch status
let statusBarItem: vscode.StatusBarItem;

// Configuration helper
function getConfig(): vscode.WorkspaceConfiguration {
	return vscode.workspace.getConfiguration('excel-power-query-editor');
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

// This method is called when your extension is activated
export function activate(context: vscode.ExtensionContext) {
	console.log('Excel Power Query Editor extension is now active!');

	// Register all commands
	const commands = [
		vscode.commands.registerCommand('excel-power-query-editor.extractFromExcel', extractFromExcel),
		vscode.commands.registerCommand('excel-power-query-editor.syncToExcel', syncToExcel),
		vscode.commands.registerCommand('excel-power-query-editor.watchFile', watchFile),
		vscode.commands.registerCommand('excel-power-query-editor.toggleWatch', toggleWatch),
		vscode.commands.registerCommand('excel-power-query-editor.stopWatching', stopWatching),
		vscode.commands.registerCommand('excel-power-query-editor.syncAndDelete', syncAndDelete),
		vscode.commands.registerCommand('excel-power-query-editor.rawExtraction', rawExtraction)
	];

	context.subscriptions.push(...commands);

	// Initialize output channel and status bar
	outputChannel = vscode.window.createOutputChannel('Excel Power Query Editor');
	updateStatusBar();
	
	log('Excel Power Query Editor extension activated');
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
		}
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to extract Power Query: ${error}`);
		console.error('Extract error:', error);
	}
}

async function syncToExcel(uri?: vscode.Uri): Promise<void> {
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

		// Read the .m file content
		const mContent = fs.readFileSync(mFile, 'utf8');
		
		// Extract just the M code (remove our comment headers)
		const mCodeMatch = mContent.match(/(?:\/\/.*\n)*\n*([\s\S]+)/);
		const cleanMCode = mCodeMatch ? mCodeMatch[1].trim() : mContent.trim();
		
		if (!cleanMCode) {
			vscode.window.showErrorMessage('No Power Query M code found in file.');
			return;
		}
		
		// Create backup of Excel file
		const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
		const backupPath = `${excelFile}.backup.${timestamp}`;
		fs.copyFileSync(excelFile, backupPath);
		
		vscode.window.showInformationMessage(`Syncing to Excel... (Backup created: ${path.basename(backupPath)})`);
		
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
				
			} catch (manualError) {
				throw new Error(`Both excel-datamashup and manual XML approaches failed. DataMashup error: ${dataMashupError}. Manual error: ${manualError}`);
			}
		}
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to sync to Excel: ${error}`);
		console.error('Sync error:', error);
		
		// If we have a backup, offer to restore it
		const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
		if (mFile) {
			const excelFile = await findExcelFile(mFile);
			if (excelFile) {
				const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
				const backupPath = `${excelFile}.backup.${timestamp}`;
				if (fs.existsSync(backupPath)) {
					const restore = await vscode.window.showErrorMessage(
						'Sync failed. Restore from backup?',
						'Restore', 'Keep Current'
					);
					if (restore === 'Restore') {
						fs.copyFileSync(backupPath, excelFile);
						vscode.window.showInformationMessage('Excel file restored from backup.');
					}
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
				log(`File changed, auto-syncing: ${path.basename(mFile)}`);
				await syncToExcel(vscode.Uri.file(mFile));
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

// This method is called when your extension is deactivated
export function deactivate() {
	// Close all file watchers
	for (const [, watcher] of fileWatchers) {
		watcher.close();
	}
	fileWatchers.clear();
}
