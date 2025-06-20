// The module 'vscode' contains the VS Code extensibility API
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { watch, FSWatcher } from 'chokidar';

// File watchers storage
const fileWatchers = new Map<string, FSWatcher>();

// This method is called when your extension is activated
export function activate(context: vscode.ExtensionContext) {
	console.log('Excel Power Query Editor extension is now active!');

	// Register all commands
	const commands = [
		vscode.commands.registerCommand('excel-power-query-editor.extractFromExcel', extractFromExcel),
		vscode.commands.registerCommand('excel-power-query-editor.syncToExcel', syncToExcel),
		vscode.commands.registerCommand('excel-power-query-editor.watchFile', watchFile),
		vscode.commands.registerCommand('excel-power-query-editor.stopWatching', stopWatching),
		vscode.commands.registerCommand('excel-power-query-editor.syncAndDelete', syncAndDelete),
		vscode.commands.registerCommand('excel-power-query-editor.rawExtraction', rawExtraction)
	];

	context.subscriptions.push(...commands);
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
		const excelFile = await findExcelFile(mFile);
		if (!excelFile) {
			vscode.window.showErrorMessage('Could not find corresponding Excel file. Please select one.');
			const selected = await selectExcelFile();
			if (!selected) {
				return;
			}
			// TODO: Implement sync with selected file
			return;
		}

		// Read the .m file content
		const content = fs.readFileSync(mFile, 'utf8');
		
		// Create backup of Excel file
		const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
		const backupPath = `${excelFile}.backup.${timestamp}`;
		fs.copyFileSync(excelFile, backupPath);
		
		vscode.window.showInformationMessage(`Syncing to Excel... (Backup created: ${path.basename(backupPath)})`);
		
		// TODO: Implement actual sync back to Excel
		// This would require writing back to the Excel DataMashup
		vscode.window.showWarningMessage('Sync to Excel functionality is not yet implemented. Backup created for safety.');
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to sync to Excel: ${error}`);
		console.error('Sync error:', error);
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

		const watcher = watch(mFile, { ignoreInitial: true });
		watcher.on('change', () => {
			vscode.window.showInformationMessage(`File changed, syncing: ${path.basename(mFile)}`);
			syncToExcel(vscode.Uri.file(mFile));
		});

		fileWatchers.set(mFile, watcher);
		vscode.window.showInformationMessage(`Now watching: ${path.basename(mFile)}`);
		
	} catch (error) {
		vscode.window.showErrorMessage(`Failed to watch file: ${error}`);
		console.error('Watch error:', error);
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
	} else {
		vscode.window.showInformationMessage(`File was not being watched: ${path.basename(mFile)}`);
	}
}

async function syncAndDelete(uri?: vscode.Uri): Promise<void> {
	const mFile = uri?.fsPath || vscode.window.activeTextEditor?.document.fileName;
	if (!mFile) {
		return;
	}

	const confirmation = await vscode.window.showWarningMessage(
		`Sync to Excel and delete ${path.basename(mFile)}?`,
		'Yes', 'No'
	);
	
	if (confirmation === 'Yes') {
		await syncToExcel(uri);
		try {
			fs.unlinkSync(mFile);
			vscode.window.showInformationMessage(`Deleted: ${path.basename(mFile)}`);
		} catch (error) {
			vscode.window.showErrorMessage(`Failed to delete file: ${error}`);
		}
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
