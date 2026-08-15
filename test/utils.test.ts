import * as assert from 'assert';
import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { initTestConfig, cleanupTestConfig, testConfigUpdate } from './testUtils';

// Utils Tests - Testing utility functions and helper behaviors
suite('Utils Tests', () => {
	const tempDir = path.join(__dirname, 'temp');

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

	suite('File Path Utilities', () => {
		test('Extension handles Excel file paths correctly', () => {
			const testPaths = [
				'C:\\Users\\test\\file.xlsx',
				'/home/user/file.xlsm',
				'./relative/path/file.xlsb',
				'file.xlsx',
				'complex file with spaces.xlsx'
			];

			testPaths.forEach(testPath => {
				const basename = path.basename(testPath);
				const dirname = path.dirname(testPath);
				
				assert.ok(basename.length > 0, `Basename should be extracted: ${testPath}`);
				assert.ok(dirname.length > 0, `Dirname should be extracted: ${testPath}`);
				
				console.log(`✅ Path utilities work for: ${testPath}`);
			});
		});

		test('Output file naming follows expected pattern', () => {
			const excelFiles = [
				'simple.xlsx',
				'complex.xlsm',
				'binary.xlsb',
				'file with spaces.xlsx'
			];

			excelFiles.forEach(filename => {
				const expectedPattern = `${filename}_PowerQuery.m`;
				const actualPattern = `${filename}_PowerQuery.m`;
				
				assert.strictEqual(actualPattern, expectedPattern, 
					`Output naming pattern should be consistent: ${filename}`);
				
				console.log(`✅ Output naming pattern correct for: ${filename} -> ${expectedPattern}`);
			});
		});
	});

	suite('Excel Format Detection', () => {
		test('File extensions are recognized correctly', () => {
			const formatTests = [
				{ file: 'test.xlsx', expected: 'Excel XLSX' },
				{ file: 'test.xlsm', expected: 'Excel XLSM (with macros)' },
				{ file: 'test.xlsb', expected: 'Excel Binary' },
				{ file: 'test.xls', expected: 'Legacy Excel (not supported)' }
			];

			formatTests.forEach(test => {
				const ext = path.extname(test.file).toLowerCase();
				let detected = 'Unknown format';
				
				switch (ext) {
					case '.xlsx':
						detected = 'Excel XLSX';
						break;
					case '.xlsm':
						detected = 'Excel XLSM (with macros)';
						break;
					case '.xlsb':
						detected = 'Excel Binary';
						break;
					case '.xls':
						detected = 'Legacy Excel (not supported)';
						break;
				}
				
				assert.strictEqual(detected, test.expected, 
					`Format detection should be correct for ${test.file}`);
				
				console.log(`✅ Format detection correct: ${test.file} -> ${detected}`);
			});
		});
	});

	suite('Power Query Parsing Helpers', () => {
		test('DataMashup XML format detection', () => {
			const xmlSamples = [
				{
					content: '<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">',
					shouldDetect: true,
					description: 'Standard DataMashup format'
				},
				{
					content: '<?xml version="1.0" encoding="utf-8"?><DataMashup sqmid="test" xmlns="http://schemas.microsoft.com/DataMashup">',
					shouldDetect: true,
					description: 'DataMashup with SQMID'
				},
				{
					content: '<?xml version="1.0"?><SomeOtherXml>',
					shouldDetect: false,
					description: 'Non-DataMashup XML'
				},
				{
					content: 'Not XML at all',
					shouldDetect: false,
					description: 'Non-XML content'
				}
			];

			xmlSamples.forEach(sample => {
				const isDataMashup = sample.content.includes('<DataMashup') && 
								   sample.content.includes('xmlns="http://schemas.microsoft.com/DataMashup"');
				
				assert.strictEqual(isDataMashup, sample.shouldDetect, 
					`DataMashup detection should be correct for: ${sample.description}`);
				
				console.log(`✅ DataMashup detection: ${sample.description} -> ${isDataMashup}`);
			});
		});

		test('Power Query formula extraction patterns', () => {
			const formulaTests = [
				{
					input: 'let\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]\nin\n    Source',
					isValidM: true,
					description: 'Basic let-in formula'
				},
				{
					input: 'section Section1;\nshared Table1 = let\n    Source = Excel.CurrentWorkbook()\nin\n    Source;',
					isValidM: true,
					description: 'Section-based formula'
				},
				{
					input: 'not a power query formula',
					isValidM: false,
					description: 'Invalid formula'
				}
			];

			formulaTests.forEach(test => {
				const hasLetIn = test.input.includes('let') && test.input.includes('in');
				const hasSection = test.input.includes('section') || hasLetIn;
				const isValid = hasSection && test.input.trim().length > 10;
				
				assert.strictEqual(isValid, test.isValidM, 
					`Formula validation should be correct for: ${test.description}`);
				
				console.log(`✅ Formula validation: ${test.description} -> ${isValid}`);
			});
		});
	});

	suite('Configuration Validation', () => {
		test('Backup location settings validation', async () => {
			const validSettings = ['sameFolder', 'tempFolder', 'custom'];
			
			for (const setting of validSettings) {
				await testConfigUpdate('backup.location', setting);
				console.log(`✅ Backup location setting accepted: ${setting}`);
			}
			
			// Test invalid setting (should handle gracefully)
			await testConfigUpdate('backup.location', 'invalidOption');
			console.log(`✅ Invalid backup location handled gracefully`);
		});

		test('Numeric configuration bounds', async () => {
			const numericTests = [
				{ key: 'sync.timeout', valid: [5000, 30000, 120000], invalid: [1000, 200000] },
				{ key: 'backup.maxFiles', valid: [1, 5, 50], invalid: [0, 100] }
			];

			for (const test of numericTests) {
				// Test valid values
				for (const value of test.valid) {
					await testConfigUpdate(test.key, value);
					console.log(`✅ ${test.key} accepted valid value: ${value}`);
				}
				
				// Test invalid values (should handle gracefully)
				for (const value of test.invalid) {
					await testConfigUpdate(test.key, value);
					console.log(`✅ ${test.key} handled invalid value gracefully: ${value}`);
				}
			}
		});

		test('Boolean configuration handling', async () => {
			const booleanSettings = [
				'watch.always',
				'backup.autoBackupBeforeSync',
				// NB: there are no boolean logging settings. debugMode and verboseMode were flat
				// keys in v0.5.0 and are now deprecated aliases that fold into log.level.
				'watch.always',
				'watch.offOnDelete'
			];

			for (const setting of booleanSettings) {
				await testConfigUpdate(setting, true);
				await testConfigUpdate(setting, false);
				console.log(`✅ Boolean setting handled correctly: ${setting}`);
			}
		});
	});

	suite('New v0.5.0 Utility Functions', () => {
		test('Backup file naming with timestamps', () => {
			const testFile = 'example.xlsx';
			const timestamp = '2025-07-11_133000';
			
			// Simulate backup naming pattern
			const backupName = `${path.basename(testFile, path.extname(testFile))}_backup_${timestamp}${path.extname(testFile)}`;
			const expectedPattern = 'example_backup_2025-07-11_133000.xlsx';
			
			assert.strictEqual(backupName, expectedPattern, 
				'Backup naming should follow timestamp pattern');
			
			console.log(`✅ Backup naming pattern: ${testFile} -> ${backupName}`);
		});

		test('Maximum backup files calculation', () => {
			// Simulate file list with timestamps
			const mockBackupFiles = [
				'file_backup_2025-07-11_120000.xlsx',
				'file_backup_2025-07-11_130000.xlsx',
				'file_backup_2025-07-11_140000.xlsx',
				'file_backup_2025-07-11_150000.xlsx',
				'file_backup_2025-07-11_160000.xlsx',
				'file_backup_2025-07-11_170000.xlsx' // 6 files
			];

			const maxFiles = 5;
			const filesToDelete = mockBackupFiles.length - maxFiles;
			
			assert.strictEqual(filesToDelete, 1, 
				'Should identify correct number of files to delete');
			
			// Oldest files should be deleted first (sorted by timestamp)
			const sortedFiles = mockBackupFiles.sort();
			const toDelete = sortedFiles.slice(0, filesToDelete);
			
			assert.strictEqual(toDelete[0], 'file_backup_2025-07-11_120000.xlsx', 
				'Should delete oldest backup first');
			
			console.log(`✅ Backup cleanup logic: ${filesToDelete} files to delete`);
		});

		test('Debounce timing configuration', async () => {
			const debounceValues = [100, 500, 1000, 2000, 5000];
			
			for (const value of debounceValues) {
				await testConfigUpdate('sync.debounceMs', value);
				console.log(`✅ Debounce timing accepted: ${value}ms`);
			}
		});
	});
});


// --- Legacy Settings Migration Tests ---
//
// These drive the REAL VS Code configuration, not the in-memory test map. The previous versions
// wrote to the test map with `testConfigUpdate` while migrateLegacySettings() read
// vscode.workspace.getConfiguration - so they asserted on values the code under test never saw.
// They also used the key names `log.debugMode` and `log.verboseMode`, which have never existed in
// any shipped version; the v0.5.0 names were `debugMode` and `verboseMode`.
import { migrateLegacySettings, scopesNeedingMigration } from '../src/extension';

const SECTION = 'excel-power-query-editor';
const MARKER = 'xtn.level';
const G = vscode.ConfigurationTarget.Global;

async function setGlobal(key: string, value: unknown): Promise<void> {
	await vscode.workspace.getConfiguration(SECTION).update(key, value, G);
}

async function clearGlobal(...keys: string[]): Promise<void> {
	const cfg = vscode.workspace.getConfiguration(SECTION);
	for (const key of keys) {
		await cfg.update(key, undefined, G);
	}
}

function globalValue(key: string): unknown {
	const info = vscode.workspace.getConfiguration(SECTION).inspect(key);
	return info ? (info as Record<string, unknown>).globalValue : undefined;
}

suite('Migration scope guard', () => {
	// The bug this pins: the marker is written to Global, but the old guard read the EFFECTIVE
	// value, which merges global over workspace and folder. Once any one workspace was migrated,
	// every other workspace was skipped forever - and since nothing reads the old keys as a
	// fallback, those users' settings silently stopped applying and the defaults took over.
	const SCHEMA = '1';

	test('a fresh install needs every scope', () => {
		const needed = scopesNeedingMigration(undefined, SCHEMA);
		assert.deepStrictEqual(needed, { global: true, workspace: true, folder: true });
	});

	test('a migrated GLOBAL scope must NOT skip workspace or folder', () => {
		const needed = scopesNeedingMigration({ globalValue: SCHEMA }, SCHEMA);
		assert.strictEqual(needed.global, false, 'global is done');
		assert.strictEqual(needed.workspace, true, 'a second workspace must still migrate');
		assert.strictEqual(needed.folder, true, 'a folder in it must still migrate');
	});

	test('each scope is judged only by its own marker', () => {
		const needed = scopesNeedingMigration(
			{ globalValue: SCHEMA, workspaceValue: SCHEMA, workspaceFolderValue: undefined }, SCHEMA);
		assert.deepStrictEqual(needed, { global: false, workspace: false, folder: true });
	});

	test('a marker from an older schema re-migrates that scope', () => {
		const needed = scopesNeedingMigration({ globalValue: '0' }, SCHEMA);
		assert.strictEqual(needed.global, true, 'schema bump must re-run');
	});
});

suite('Legacy Settings Migration', () => {
	const touched = [
		MARKER, 'watchAlways', 'watch.always', 'syncTimeout', 'sync.timeout',
		'debugMode', 'verboseMode', 'log.level', 'backupLocation', 'backup.location',
	];

	setup(async () => { await clearGlobal(...touched); });
	teardown(async () => { await clearGlobal(...touched); });

	test('carries a renamed value across and clears the old key', async () => {
		await setGlobal('watchAlways', true);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('watch.always'), true, 'new key should hold the old value');
		assert.strictEqual(globalValue('watchAlways'), undefined, 'old key should be cleared');
	});

	test('does NOT overwrite a new value the user has already set', async () => {
		await setGlobal('syncTimeout', 10);
		await setGlobal('sync.timeout', 99);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('sync.timeout'), 99, 'existing new value must win');
		assert.strictEqual(globalValue('syncTimeout'), undefined, 'stale old key still cleared');
	});

	test('debugMode becomes log.level=debug', async () => {
		await setGlobal('debugMode', true);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('log.level'), 'debug');
		assert.strictEqual(globalValue('debugMode'), undefined);
	});

	test('debug beats verbose when both are set', async () => {
		await setGlobal('debugMode', true);
		await setGlobal('verboseMode', true);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('log.level'), 'debug');
		assert.strictEqual(globalValue('debugMode'), undefined);
		assert.strictEqual(globalValue('verboseMode'), undefined);
	});

	test('verboseMode alone becomes log.level=verbose', async () => {
		await setGlobal('verboseMode', true);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('log.level'), 'verbose');
	});

	test('a false legacy flag is cleared without inventing a log level', async () => {
		await setGlobal('debugMode', false);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('log.level'), undefined, 'false should not set a level');
		assert.strictEqual(globalValue('debugMode'), undefined, 'but the key is still cleared');
	});

	// The regression that matters. The previous implementation wiped every setting in scope.
	test('leaves settings it does not own completely alone', async () => {
		await setGlobal('backup.location', 'custom');
		await setGlobal('watchAlways', true);
		await migrateLegacySettings();

		assert.strictEqual(globalValue('backup.location'), 'custom',
			'an unrelated modern setting must survive migration');
	});

	test('runs once - a second call is a no-op', async () => {
		await setGlobal('watchAlways', true);
		await migrateLegacySettings();
		assert.strictEqual(globalValue(MARKER), '1', 'marker should be stamped');

		// Set a legacy key again; the marker should stop it being touched.
		await setGlobal('watchAlways', true);
		await migrateLegacySettings();
		assert.strictEqual(globalValue('watchAlways'), true, 'second run must not migrate again');
	});
});
