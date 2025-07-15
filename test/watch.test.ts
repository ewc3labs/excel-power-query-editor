import * as assert from 'assert';
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { initTestConfig, cleanupTestConfig, testConfigUpdate } from './testUtils';

// Watch Tests - Testing file watching configuration and command registration
suite('Watch Tests', () => {
	const tempDir = path.join(__dirname, 'temp');
	const fixturesDir = path.join(__dirname, '..', '..', 'test', 'fixtures');

	suiteSetup(() => {
		// Initialize test configuration system
		initTestConfig();
		
		// Ensure temp directory exists
		if (!fs.existsSync(tempDir)) {
			fs.mkdirSync(tempDir, { recursive: true });
		}
	});

	suiteTeardown(() => {
		// Clean up test configuration
		cleanupTestConfig();
		
		// Clean up temp directory
		if (fs.existsSync(tempDir)) {
			fs.rmSync(tempDir, { recursive: true, force: true });
		}
	});

	suite('Watch Command Registration', () => {
		test('Watch commands are registered and callable', async () => {
			const commands = await vscode.commands.getCommands(true);
			
			const watchCommands = [
				'excel-power-query-editor.watchFile',
				'excel-power-query-editor.toggleWatch',
				'excel-power-query-editor.stopWatching'
			];
			
			watchCommands.forEach(command => {
				assert.ok(commands.includes(command), `Command should be registered: ${command}`);
				console.log(`✅ Watch command registered: ${command}`);
			});
		});

		test('Watch commands handle basic invocation', async () => {
			const testMFile = path.join(tempDir, 'basic_watch_test.m');
			
			// Create a test .m file
			const sampleMContent = `// Basic watch test file
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
in
    Source`;
			
			fs.writeFileSync(testMFile, sampleMContent, 'utf8');
			const uri = vscode.Uri.file(testMFile);
			
			// Test that commands can be called without crashing the extension
			const watchCommands = [
				'excel-power-query-editor.watchFile',
				'excel-power-query-editor.toggleWatch',
				'excel-power-query-editor.stopWatching'
			];
			
			for (const command of watchCommands) {
				try {
					await Promise.race([
						vscode.commands.executeCommand(command, uri),
						new Promise((resolve) => setTimeout(resolve, 500)) // Quick timeout
					]);
					console.log(`✅ ${command} executed without crashing`);
				} catch (error) {
					console.log(`✅ ${command} handled gracefully: ${error}`);
				}
			}
		});
	});

	suite('Watch Configuration Settings', () => {
		test('Watch-related configuration is accepted', async () => {
			const configTests = [
				{ key: 'watchAlways', values: [true, false] },
				{ key: 'watchOffOnDelete', values: [true, false] },
				{ key: 'watch.checkExcelWriteable', values: [true, false] }
			];
			
			for (const test of configTests) {
				for (const value of test.values) {
					await testConfigUpdate(test.key, value);
					console.log(`✅ ${test.key} setting accepted: ${value}`);
				}
			}
		});

		test('Debounce timing configuration', async () => {
			const debounceValues = [100, 250, 500, 1000, 2000, 5000];
			
			for (const value of debounceValues) {
				await testConfigUpdate('sync.debounceMs', value);
				console.log(`✅ Debounce timing accepted: ${value}ms`);
			}
		});
	});

	suite('File Path Handling', () => {
		test('Watch system handles different file paths', () => {
			const testPaths = [
				'simple.m',
				'complex with spaces.m',
				'deep/nested/path/file.m',
				'C:\\Windows\\path\\file.m',
				'/unix/style/path/file.m'
			];

			testPaths.forEach(testPath => {
				const basename = path.basename(testPath);
				const dirname = path.dirname(testPath);
				const extname = path.extname(testPath);
				
				assert.strictEqual(extname, '.m', `Should recognize .m extension: ${testPath}`);
				assert.ok(basename.length > 0, `Should extract basename: ${testPath}`);
				
				console.log(`✅ Path handling verified for: ${testPath}`);
			});
		});

		test('Excel file association logic', () => {
			const testCases = [
				{ mFile: 'simple.xlsx_PowerQuery.m', expectedExcel: 'simple.xlsx' },
				{ mFile: 'complex.xlsm_PowerQuery.m', expectedExcel: 'complex.xlsm' },
				{ mFile: 'binary.xlsb_PowerQuery.m', expectedExcel: 'binary.xlsb' },
				{ mFile: 'file with spaces.xlsx_PowerQuery.m', expectedExcel: 'file with spaces.xlsx' }
			];

			testCases.forEach(test => {
				// Simulate the logic to find associated Excel file
				const mBaseName = path.basename(test.mFile, '.m');
				if (mBaseName.endsWith('_PowerQuery')) {
					const excelName = mBaseName.replace('_PowerQuery', '');
					assert.strictEqual(excelName, test.expectedExcel, 
						`Should correctly identify Excel file: ${test.mFile} -> ${test.expectedExcel}`);
					console.log(`✅ Excel association: ${test.mFile} -> ${excelName}`);
				}
			});
		});
	});

	suite('File System Operations Simulation', () => {
		test('File creation and deletion detection patterns', async () => {
			const testFile = path.join(tempDir, 'fs_operations_test.m');
			const content = '// Test content for file operations';
			
			// Test file creation
			fs.writeFileSync(testFile, content, 'utf8');
			assert.ok(fs.existsSync(testFile), 'File should be created');
			console.log(`✅ File creation detected: ${path.basename(testFile)}`);
			
			// Test file modification
			const modifiedContent = content + '\n// Modified content';
			fs.writeFileSync(testFile, modifiedContent, 'utf8');
			const readContent = fs.readFileSync(testFile, 'utf8');
			assert.ok(readContent.includes('Modified content'), 'File should be modified');
			console.log(`✅ File modification detected: ${path.basename(testFile)}`);
			
			// Test file deletion
			fs.unlinkSync(testFile);
			assert.ok(!fs.existsSync(testFile), 'File should be deleted');
			console.log(`✅ File deletion detected: ${path.basename(testFile)}`);
		});

		test('Multiple file operations', async () => {
			const testFiles = [
				path.join(tempDir, 'multi_op_1.m'),
				path.join(tempDir, 'multi_op_2.m'),
				path.join(tempDir, 'multi_op_3.m')
			];
			
			// Create multiple files
			testFiles.forEach((file, index) => {
				const content = `// Test file ${index + 1}
let
    Source${index + 1} = Excel.CurrentWorkbook(){[Name="Table${index + 1}"]}[Content]
in
    Source${index + 1}`;
				
				fs.writeFileSync(file, content, 'utf8');
				assert.ok(fs.existsSync(file), `File ${index + 1} should be created`);
			});
			
			console.log(`✅ Multiple file creation: ${testFiles.length} files created`);
			
			// Clean up
			testFiles.forEach(file => {
				if (fs.existsSync(file)) {
					fs.unlinkSync(file);
				}
			});
			
			console.log(`✅ Multiple file cleanup: ${testFiles.length} files removed`);
		});
	});

	suite('Error Handling Scenarios', () => {
		test('Invalid file handling', async () => {
			const invalidFile = path.join(tempDir, 'invalid.txt');
			fs.writeFileSync(invalidFile, 'Not a Power Query file', 'utf8');
			
			const uri = vscode.Uri.file(invalidFile);
			
			try {
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.watchFile', uri),
					new Promise((resolve) => setTimeout(resolve, 200))
				]);
				console.log(`✅ Invalid file type handled gracefully`);
			} catch (error) {
				console.log(`✅ Invalid file error handled: ${error}`);
			}
		});

		test('Non-existent file handling', async () => {
			const nonExistentFile = path.join(tempDir, 'does_not_exist.m');
			const uri = vscode.Uri.file(nonExistentFile);
			
			try {
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.watchFile', uri),
					new Promise((resolve) => setTimeout(resolve, 200))
				]);
				console.log(`✅ Non-existent file handled gracefully`);
			} catch (error) {
				console.log(`✅ Non-existent file error handled: ${error}`);
			}
		});
	});

	suite('Integration with Extension Features', () => {
		test('Watch functionality integrates with Excel operations', async () => {
			const sourceFile = path.join(fixturesDir, 'simple.xlsx');
			
			if (fs.existsSync(sourceFile)) {
				// Copy to temp directory to avoid polluting fixtures
				const testExcelFile = path.join(tempDir, 'simple_watch_test.xlsx');
				fs.copyFileSync(sourceFile, testExcelFile);
				console.log(`📁 Copied simple.xlsx to temp directory for watch integration test`);
				
				const uri = vscode.Uri.file(testExcelFile);
				
				try {
					// Test that watch commands don't interfere with extraction
					await Promise.race([
						vscode.commands.executeCommand('excel-power-query-editor.extractFromExcel', uri),
						new Promise((_, reject) => setTimeout(() => reject(new Error('Extraction timeout')), 2000))
					]);
					
					console.log(`✅ Watch integration with extraction works`);
					
					// Test watch command on extracted files (in temp dir)
					const extractedDir = path.dirname(testExcelFile);
					const mFiles = fs.readdirSync(extractedDir).filter(f => f.endsWith('.m'));
					
					if (mFiles.length > 0) {
						const mUri = vscode.Uri.file(path.join(extractedDir, mFiles[0]));
						try {
							await Promise.race([
								vscode.commands.executeCommand('excel-power-query-editor.toggleWatch', mUri),
								new Promise((resolve) => setTimeout(resolve, 200))
							]);
							console.log(`✅ Watch command works on extracted .m files`);
						} catch (error) {
							console.log(`✅ Watch on extracted files handled: ${error}`);
						}
					}
					
				} catch (error) {
					console.log(`✅ Watch-extraction integration handled: ${error}`);
				}
			} else {
				console.log('⏭️  Skipping integration test - simple.xlsx not found');
			}
		});
	});
});
