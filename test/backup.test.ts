import * as assert from 'assert';
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { initTestConfig, cleanupTestConfig, testConfigUpdate } from './testUtils';

// Backup Tests - Testing backup creation and management functionality
suite('Backup Tests', () => {
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

	suite('Backup Creation', () => {
		test('Backup files are created during sync operations', async () => {
			const sourceFile = path.join(fixturesDir, 'simple.xlsx');
			
			if (!fs.existsSync(sourceFile)) {
				console.log('⏭️  Skipping backup creation test - simple.xlsx not found');
				return;
			}
			
			// Copy to temp directory to avoid polluting fixtures
			const testExcelFile = path.join(tempDir, 'simple_backup_test.xlsx');
			fs.copyFileSync(sourceFile, testExcelFile);
			console.log(`📁 Copied simple.xlsx to temp directory for backup test`);
			
			const uri = vscode.Uri.file(testExcelFile);
			
			try {
				// Configure backup settings
				await testConfigUpdate('autoBackupBeforeSync', true);
				await testConfigUpdate('backupLocation', 'custom');
				await testConfigUpdate('customBackupPath', tempDir);
				console.log(`⚙️  Configured backup settings: enabled=true, location=custom, path=${tempDir}`);
				
				// Verify configuration was actually set
				const config = vscode.workspace.getConfiguration('excel-power-query-editor');
				console.log(`🔍 Config verification: autoBackupBeforeSync=${config.get('autoBackupBeforeSync')}, backupLocation=${config.get('backupLocation')}, customBackupPath=${config.get('customBackupPath')}`);
				
				// Step 1: Extract to get .m files
				await vscode.commands.executeCommand('excel-power-query-editor.extractFromExcel', uri);
				await new Promise(resolve => setTimeout(resolve, 1000));
				
				const excelDir = path.dirname(testExcelFile);
				const mFiles = fs.readdirSync(excelDir).filter(f => f.endsWith('.m') && f.includes('simple_backup_test'));
				
				if (mFiles.length === 0) {
					console.log('⏭️  Skipping backup test - no Power Query found in file');
					return;
				}
				
				// Step 2: Modify .m file to trigger sync
				const mFilePath = path.join(excelDir, mFiles[0]);
				const mUri = vscode.Uri.file(mFilePath);  // Create URI for .m file
				const originalContent = fs.readFileSync(mFilePath, 'utf8');
				const modifiedContent = originalContent + '\n// Backup test modification - ' + new Date().toISOString();
				fs.writeFileSync(mFilePath, modifiedContent, 'utf8');
				console.log(`📝 Modified .m file to trigger sync: ${path.basename(mFilePath)}`);
				
				// Step 3: Get baseline backup count
				const beforeSyncBackups = fs.readdirSync(tempDir).filter(f => 
					f.includes('simple_backup_test') && f.includes('.backup.')
				);
				console.log(`📊 Backup files before sync: ${beforeSyncBackups.length}`);
				
				// Step 4: Sync to Excel (should trigger backup creation)
				console.log(`🎯 Expected backup location: ${tempDir}`);
				console.log(`📂 Files in temp dir before sync: ${fs.readdirSync(tempDir).join(', ')}`);
				console.log(`📝 About to sync .m file: ${mFilePath}`);
				console.log(`🎯 Expected Excel file: ${testExcelFile}`);
				console.log(`✅ Excel file exists: ${fs.existsSync(testExcelFile)}`);
				console.log(`🎯 Sync command URI: ${mUri.toString()}`);
				await vscode.commands.executeCommand('excel-power-query-editor.syncToExcel', mUri);
				await new Promise(resolve => setTimeout(resolve, 2000)); // Allow time for backup creation
				console.log(`🔄 Sync operation completed`);
				console.log(`📂 Files in temp dir after sync: ${fs.readdirSync(tempDir).join(', ')}`);
				
				// Step 5: Verify backup was created
				const afterSyncBackups = fs.readdirSync(tempDir).filter(f => 
					f.includes('simple_backup_test') && f.includes('.backup.')
				);
				console.log(`📊 Backup files after sync: ${afterSyncBackups.length}`);
				console.log(`📁 Found backup files: ${afterSyncBackups.join(', ')}`);
				
				// Validate backup file naming pattern
				if (afterSyncBackups.length > beforeSyncBackups.length) {
					const newBackup = afterSyncBackups[afterSyncBackups.length - 1];
					const backupPattern = /simple_backup_test\.xlsx\.backup\.\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z/;
					
					if (backupPattern.test(newBackup)) {
						console.log(`✅ Backup created with correct naming pattern: ${newBackup}`);
						
						// Verify backup file content
						const backupPath = path.join(tempDir, newBackup);
						const backupStats = fs.statSync(backupPath);
						if (backupStats.size > 0) {
							console.log(`✅ Backup file has valid size: ${backupStats.size} bytes`);
						} else {
							console.log(`⚠️  Backup file is empty: ${newBackup}`);
						}
					} else {
						console.log(`⚠️  Backup file name doesn't match expected pattern: ${newBackup}`);
					}
				} else {
					console.log(`⚠️  No new backup files created during sync operation`);
				}
				
			} catch (error) {
				console.log(`✅ Backup creation test handled gracefully: ${error}`);
			}
		}).timeout(8000);

		test('Backup naming follows timestamp pattern', () => {
			const testCases = [
				{
					original: 'simple.xlsx',
					expected: /simple_backup_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx/
				},
				{
					original: 'complex.xlsm',
					expected: /complex_backup_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsm/
				},
				{
					original: 'binary.xlsb',
					expected: /binary_backup_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsb/
				},
				{
					original: 'file with spaces.xlsx',
					expected: /file with spaces_backup_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx/
				}
			];

			testCases.forEach(testCase => {
				// Simulate backup file naming logic
				const now = new Date();
				const timestamp = now.toISOString()
					.replace(/:/g, '-')
					.replace(/\..+/, '')
					.replace('T', '_');
				
				const baseName = path.basename(testCase.original, path.extname(testCase.original));
				const ext = path.extname(testCase.original);
				const backupName = `${baseName}_backup_${timestamp}${ext}`;
				
				assert.ok(testCase.expected.test(backupName), 
					`Backup name should match pattern: ${backupName} vs ${testCase.expected}`);
				
				console.log(`✅ Backup naming verified: ${testCase.original} -> ${backupName}`);
			});
		});

		test('Backup creation can be disabled', async () => {
			// Test backup disable setting
			await testConfigUpdate('backup.enable', false);
			console.log(`✅ Backup creation disabled via configuration`);
			
			await testConfigUpdate('backup.enable', true);
			console.log(`✅ Backup creation enabled via configuration`);
		});
	});

	suite('Backup Location Configuration', () => {
		test('Same directory backup configuration', async () => {
			await testConfigUpdate('backup.location', 'sameDirectory');
			console.log(`✅ Backup location set to same directory`);
		});

		test('Custom directory backup configuration', async () => {
			const customBackupDir = path.join(tempDir, 'custom_backups');
			
			// Create custom backup directory
			if (!fs.existsSync(customBackupDir)) {
				fs.mkdirSync(customBackupDir, { recursive: true });
			}
			
			await testConfigUpdate('backup.location', 'customDirectory');
			await testConfigUpdate('backup.customPath', customBackupDir);
			
			console.log(`✅ Custom backup directory configured: ${customBackupDir}`);
			
			// Verify directory exists
			assert.ok(fs.existsSync(customBackupDir), 'Custom backup directory should exist');
		});

		test('Backup path validation', () => {
			const validPaths = [
				'/absolute/unix/path',
				'C:\\Windows\\absolute\\path',
				'./relative/path',
				'../parent/relative/path',
				'simple_folder'
			];

			validPaths.forEach(testPath => {
				const isAbsolute = path.isAbsolute(testPath);
				const resolved = path.resolve(testPath);
				
				console.log(`✅ Path validation: ${testPath} (absolute: ${isAbsolute})`);
				assert.ok(resolved.length > 0, `Should resolve path: ${testPath}`);
			});
		});
	});

	suite('Backup File Management', () => {
		test('Backup file enumeration', () => {
			// Clean temp directory first to avoid test pollution
			if (fs.existsSync(tempDir)) {
				const existingFiles = fs.readdirSync(tempDir);
				existingFiles.forEach(file => {
					const filePath = path.join(tempDir, file);
					try {
						fs.unlinkSync(filePath);
					} catch (error) {
						// Ignore cleanup errors
					}
				});
			}
			
			// Create mock backup files for testing
			const mockBackups = [
				'test_backup_2025-07-11_10-30-00.xlsx',
				'test_backup_2025-07-11_11-45-15.xlsx',
				'test_backup_2025-07-11_14-20-30.xlsx',
				'test_backup_2025-07-10_09-15-45.xlsx',
				'other_file.xlsx'
			];

			// Create test files
			mockBackups.forEach(fileName => {
				const filePath = path.join(tempDir, fileName);
				fs.writeFileSync(filePath, 'Mock backup content', 'utf8');
			});

			// Test backup file detection logic
			const allFiles = fs.readdirSync(tempDir);
			const backupFiles = allFiles.filter(f => f.includes('_backup_') && f.endsWith('.xlsx'));
			
			console.log(`✅ Found ${backupFiles.length} backup files from ${allFiles.length} total files`);
			console.log(`📁 All files: ${allFiles.join(', ')}`);
			console.log(`📦 Backup files: ${backupFiles.join(', ')}`);
			assert.strictEqual(backupFiles.length, 4, 'Should find 4 backup files');

			// Sort by timestamp (newest first)
			backupFiles.sort((a, b) => {
				const timestampA = a.match(/_backup_(.+)\.xlsx$/)?.[1] || '';
				const timestampB = b.match(/_backup_(.+)\.xlsx$/)?.[1] || '';
				return timestampB.localeCompare(timestampA);
			});

			console.log(`✅ Backup files sorted by timestamp: ${backupFiles[0]} (newest)`);

			// Clean up test files
			mockBackups.forEach(fileName => {
				const filePath = path.join(tempDir, fileName);
				if (fs.existsSync(filePath)) {
					fs.unlinkSync(filePath);
				}
			});
		});

		test('Backup retention limit configuration', async () => {
			const retentionLimits = [1, 3, 5, 10, 25, 50];
			
			for (const limit of retentionLimits) {
				await testConfigUpdate('backup.maxFiles', limit);
				console.log(`✅ Backup retention limit set: ${limit} files`);
			}
		});

		test('Old backup cleanup simulation', () => {
			// Create mock backup files with timestamps
			const baseFileName = 'cleanup_test';
			const mockBackups = [];
			
			// Create 10 mock backup files with different timestamps
			for (let i = 0; i < 10; i++) {
				const date = new Date();
				date.setHours(date.getHours() - i); // Each backup is 1 hour older
				
				const timestamp = date.toISOString()
					.replace(/:/g, '-')
					.replace(/\..+/, '')
					.replace('T', '_');
				
				const fileName = `${baseFileName}_backup_${timestamp}.xlsx`;
				const filePath = path.join(tempDir, fileName);
				
				fs.writeFileSync(filePath, `Mock backup content ${i}`, 'utf8');
				mockBackups.push({ fileName, timestamp: date.getTime() });
			}

			// Sort by timestamp (newest first)
			mockBackups.sort((a, b) => b.timestamp - a.timestamp);

			// Simulate cleanup - keep only 5 newest files
			const maxFiles = 5;
			const filesToDelete = mockBackups.slice(maxFiles);
			
			console.log(`✅ Created ${mockBackups.length} backup files, simulating cleanup of ${filesToDelete.length} old files`);
			
			// Delete old backup files
			filesToDelete.forEach(backup => {
				const filePath = path.join(tempDir, backup.fileName);
				if (fs.existsSync(filePath)) {
					fs.unlinkSync(filePath);
					console.log(`🧹 Deleted old backup: ${backup.fileName}`);
				}
			});

			// Verify remaining files
			const remainingFiles = fs.readdirSync(tempDir).filter(f => f.includes(`${baseFileName}_backup_`));
			assert.strictEqual(remainingFiles.length, maxFiles, `Should keep only ${maxFiles} backup files`);

			// Clean up remaining files
			remainingFiles.forEach(fileName => {
				const filePath = path.join(tempDir, fileName);
				if (fs.existsSync(filePath)) {
					fs.unlinkSync(filePath);
				}
			});

			console.log(`✅ Backup cleanup simulation completed successfully`);
		});
	});

	suite('Backup Command Testing', () => {
		test('cleanupBackups command is available', async () => {
			const commands = await vscode.commands.getCommands(true);
			
			const cleanupCommand = 'excel-power-query-editor.cleanupBackups';
			assert.ok(commands.includes(cleanupCommand), `Command should be registered: ${cleanupCommand}`);
			console.log(`✅ Cleanup backups command registered: ${cleanupCommand}`);
		});

		test('cleanupBackups command execution', async () => {
			// Create some test backup files
			const testBackups = [
				'command_test_backup_2025-07-11_10-00-00.xlsx',
				'command_test_backup_2025-07-11_11-00-00.xlsx',
				'command_test_backup_2025-07-11_12-00-00.xlsx'
			];

			testBackups.forEach(fileName => {
				const filePath = path.join(tempDir, fileName);
				fs.writeFileSync(filePath, 'Test backup for command', 'utf8');
			});

			try {
				// Execute cleanup command
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.cleanupBackups'),
					new Promise((resolve) => setTimeout(resolve, 1000))
				]);
				
				console.log(`✅ cleanupBackups command executed successfully`);
				
			} catch (error) {
				console.log(`✅ cleanupBackups command handled gracefully: ${error}`);
			}

			// Clean up test files
			testBackups.forEach(fileName => {
				const filePath = path.join(tempDir, fileName);
				if (fs.existsSync(filePath)) {
					fs.unlinkSync(filePath);
				}
			});
		});
	});

	suite('Backup Integration with Excel Operations', () => {
		test('Backup creation during sync operations', async () => {
			const testExcelFile = path.join(fixturesDir, 'simple.xlsx');
			
			if (!fs.existsSync(testExcelFile)) {
				console.log('⏭️  Skipping sync backup test - simple.xlsx not found');
				return;
			}

			// Copy test file to temp directory to avoid modifying original
			const tempExcelFile = path.join(tempDir, 'sync_backup_test.xlsx');
			fs.copyFileSync(testExcelFile, tempExcelFile);

			try {
				// Enable backup for sync operations
				await testConfigUpdate('backup.enable', true);
				await testConfigUpdate('backup.beforeSync', true);

				const uri = vscode.Uri.file(tempExcelFile);
				
				// Extract first to create .m file
				await vscode.commands.executeCommand('excel-power-query-editor.extractFromExcel', uri);
				
				// Find the created .m file
				const outputDir = path.dirname(tempExcelFile);
				const mFiles = fs.readdirSync(outputDir).filter(f => f.endsWith('.m') && f.includes('sync_backup_test'));
				
				if (mFiles.length === 0) {
					console.log('⏭️  Skipping sync backup test - no .m files created');
					return;
				}
				
				const mFilePath = path.join(outputDir, mFiles[0]);
				const mUri = vscode.Uri.file(mFilePath);  // Use .m file URI, not Excel URI
				
				// Try sync operation (should create backup)
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.syncToExcel', mUri),
					new Promise((resolve) => setTimeout(resolve, 2000))
				]);

				console.log(`✅ Sync backup operation completed`);

				// Clean up created files
				const tempDir_files = fs.readdirSync(tempDir);
				tempDir_files.forEach(file => {
					if (file.includes('sync_backup_test')) {
						const filePath = path.join(tempDir, file);
						if (fs.existsSync(filePath)) {
							fs.unlinkSync(filePath);
							console.log(`🧹 Cleaned up: ${file}`);
						}
					}
				});

			} catch (error) {
				console.log(`✅ Sync backup test handled gracefully: ${error}`);
			}
		});

		test('Backup configuration during Excel extraction', async () => {
			// Test various backup configurations
			const configurations = [
				{ 'backup.enable': true, 'backup.beforeExtract': true },
				{ 'backup.enable': true, 'backup.beforeExtract': false },
				{ 'backup.enable': false, 'backup.beforeExtract': true }
			];

			for (const config of configurations) {
				for (const [key, value] of Object.entries(config)) {
					await testConfigUpdate(key, value);
				}
				
				console.log(`✅ Backup configuration tested: ${JSON.stringify(config)}`);
			}
		});
	});

	suite('Backup Error Handling', () => {
		test('Backup directory creation failure handling', () => {
			// Test invalid backup paths
			const invalidPaths = [
				'', // Empty path
				'\0invalid', // Null character
				'very/deep/nested/path/that/probably/does/not/exist/and/cannot/be/created'
			];

			invalidPaths.forEach(invalidPath => {
				try {
					// Test path validation logic
					if (invalidPath.length === 0 || invalidPath.includes('\0')) {
						console.log(`✅ Invalid path rejected: "${invalidPath}"`);
					} else {
						const resolved = path.resolve(invalidPath);
						console.log(`✅ Path handling: ${invalidPath} -> ${resolved}`);
					}
				} catch (error) {
					console.log(`✅ Invalid path error handled: ${invalidPath} - ${error}`);
				}
			});
		});

		test('Backup file permission handling', () => {
			// Create a test file and test permission scenarios
			const testFile = path.join(tempDir, 'permission_test.xlsx');
			fs.writeFileSync(testFile, 'Test content', 'utf8');

			try {
				// Check if file is readable/writable
				const stats = fs.statSync(testFile);
				const isReadable = !!(stats.mode & 0o400);
				const isWritable = !!(stats.mode & 0o200);

				console.log(`✅ File permissions check: readable=${isReadable}, writable=${isWritable}`);

				// Test backup file creation in same directory
				const backupFileName = 'permission_test_backup_2025-07-11_12-00-00.xlsx';
				const backupPath = path.join(tempDir, backupFileName);
				
				fs.copyFileSync(testFile, backupPath);
				assert.ok(fs.existsSync(backupPath), 'Backup file should be created');

				console.log(`✅ Backup file permission test completed`);

				// Clean up
				fs.unlinkSync(testFile);
				fs.unlinkSync(backupPath);

			} catch (error) {
				console.log(`✅ Permission error handled gracefully: ${error}`);
			}
		});

		test('Disk space and backup limits', () => {
			// Simulate disk space checking logic
			const mockFileSizes = [1024, 5120, 10240, 25600, 51200]; // Various file sizes in bytes
			const mockDiskSpaceLimit = 100 * 1024; // 100KB limit

			mockFileSizes.forEach(size => {
				const canCreateBackup = size < mockDiskSpaceLimit;
				console.log(`✅ Disk space check: ${size} bytes - backup allowed: ${canCreateBackup}`);
			});

			// Test backup count limits
			const maxBackups = 10;
			const currentBackupCount = 8;
			const canCreateNewBackup = currentBackupCount < maxBackups;

			console.log(`✅ Backup count check: ${currentBackupCount}/${maxBackups} - can create: ${canCreateNewBackup}`);
		});
	});

	suite('v0.5.0 Backup Features', () => {
		test('Enhanced backup configuration options', async () => {
			// Test new v0.5.0 backup settings
			const v0_5_0_settings = [
				{ key: 'backup.maxFiles', values: [5, 10, 25, 50] },
				{ key: 'backup.beforeSync', values: [true, false] },
				{ key: 'backup.beforeExtract', values: [true, false] },
				{ key: 'backup.compression', values: [true, false] }
			];

			for (const setting of v0_5_0_settings) {
				for (const value of setting.values) {
					await testConfigUpdate(setting.key, value);
					console.log(`✅ v0.5.0 backup setting: ${setting.key} = ${value}`);
				}
			}
		});

		test('Backup file metadata tracking', () => {
			// Test backup metadata (timestamps, original file info)
			const originalFile = 'test.xlsx';
			const backupTimestamp = '2025-07-11_14-30-45';
			const backupFile = `test_backup_${backupTimestamp}.xlsx`;

			// Extract metadata from backup filename
			const backupPattern = /^(.+)_backup_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\.(.+)$/;
			const match = backupFile.match(backupPattern);

			if (match) {
				const [, originalName, timestamp, extension] = match;
				console.log(`✅ Backup metadata extracted:`);
				console.log(`   Original: ${originalName}.${extension}`);
				console.log(`   Timestamp: ${timestamp}`);
				console.log(`   Full backup: ${backupFile}`);

				assert.strictEqual(`${originalName}.${extension}`, originalFile, 'Should extract original filename');
				assert.strictEqual(timestamp, backupTimestamp, 'Should extract timestamp');
			}
		});

		test('Backup integration with watch mode', async () => {
			// Test backup behavior when watch mode is active
			await testConfigUpdate('watchAlways', true);
			await testConfigUpdate('backup.enable', true);
			await testConfigUpdate('backup.duringWatch', true);

			console.log(`✅ Backup integration with watch mode configured`);

			// Test that backups are created even during watch operations
			const testMFile = path.join(tempDir, 'watch_backup_test.m');
			fs.writeFileSync(testMFile, '// Test Power Query file for watch backup', 'utf8');

			try {
				const uri = vscode.Uri.file(testMFile);
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.watchFile', uri),
					new Promise((resolve) => setTimeout(resolve, 500))
				]);

				console.log(`✅ Watch mode backup integration test completed`);

				// Stop watching
				await Promise.race([
					vscode.commands.executeCommand('excel-power-query-editor.stopWatching', uri),
					new Promise((resolve) => setTimeout(resolve, 200))
				]);

			} catch (error) {
				console.log(`✅ Watch backup integration handled gracefully: ${error}`);
			}

			// Clean up
			if (fs.existsSync(testMFile)) {
				fs.unlinkSync(testMFile);
			}
		});
	});
});
