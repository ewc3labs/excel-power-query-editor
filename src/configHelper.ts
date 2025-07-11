import * as vscode from 'vscode';

/**
 * Test-aware configuration interface that works in both extension and test environments
 */
export interface ConfigHelper {
    get<T>(section: string, defaultValue?: T): T | undefined;
    has(section: string): boolean;
    update?(section: string, value: any, configurationTarget?: any): Thenable<void>;
}

// Global test configuration store - populated during tests
let testConfig: Map<string, any> | null = null;

/**
 * Set test configuration (called from test setup)
 */
export function setTestConfig(config: Map<string, any> | null): void {
    testConfig = config;
}

/**
 * Check if we're running in test environment
 */
function isTestEnvironment(): boolean {
    return testConfig !== null;
}

/**
 * Get configuration - automatically uses test config in test environment,
 * real VS Code config in extension environment
 */
export function getConfig(): ConfigHelper {
    if (isTestEnvironment() && testConfig) {
        // Return test configuration
        return {
            get<T>(section: string, defaultValue?: T): T | undefined {
                return testConfig!.get(section) ?? defaultValue;
            },
            has(section: string): boolean {
                return testConfig!.has(section);
            }
        };
    } else {
        // Return real VS Code configuration
        const realConfig = vscode.workspace.getConfiguration('excel-power-query-editor');
        return {
            get<T>(section: string, defaultValue?: T): T | undefined {
                return realConfig.get(section, defaultValue);
            },
            has(section: string): boolean {
                return realConfig.has(section);
            },
            update: realConfig.update.bind(realConfig)
        };
    }
}

/**
 * Default configuration values for the extension
 */
export const DEFAULT_CONFIG = {
    outputDirectory: '',
    backupFolder: '',
    createBackups: false,
    backupLocation: 'sameFolder',
    customBackupPath: '',
    autoWatch: false,
    verbose: false,
    maxBackups: 10,
    compressionLevel: 6,
    queryNaming: 'descriptive',
    autoCleanupBackups: true,
    'backup.maxFiles': 5
};
