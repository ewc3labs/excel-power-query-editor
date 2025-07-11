import * as vscode from 'vscode';
import { setTestConfig, DEFAULT_CONFIG } from '../src/configHelper';

/**
 * Test configuration store
 */
let testConfigStore = new Map<string, any>();
let originalGetConfiguration: typeof vscode.workspace.getConfiguration | null = null;

/**
 * Initialize test configuration with defaults and mock VS Code config system
 */
export function initTestConfig(): void {
    testConfigStore = new Map(Object.entries(DEFAULT_CONFIG));
    setTestConfig(testConfigStore);
    
    // Backup original VS Code config function
    if (!originalGetConfiguration) {
        originalGetConfiguration = vscode.workspace.getConfiguration;
    }
    
    // Mock VS Code configuration system universally
    vscode.workspace.getConfiguration = ((section?: string) => {
        const config = {
            get: <T>(key: string, defaultValue?: T): T => {
                const fullKey = section ? `${section}.${key}` : key;
                const value = testConfigStore.get(fullKey) ?? testConfigStore.get(key);
                return (value !== undefined ? value : defaultValue) as T;
            },
            update: async (key: string, value: any) => {
                const fullKey = section ? `${section}.${key}` : key;
                testConfigStore.set(fullKey, value);
                testConfigStore.set(key, value); // Also set without section for compatibility
                return Promise.resolve();
            },
            has: (key: string): boolean => {
                const fullKey = section ? `${section}.${key}` : key;
                return testConfigStore.has(fullKey) || testConfigStore.has(key);
            },
            inspect: (key: string) => {
                const fullKey = section ? `${section}.${key}` : key;
                const value = testConfigStore.get(fullKey) ?? testConfigStore.get(key);
                const defaultValue = (DEFAULT_CONFIG as Record<string, any>)[key];
                return {
                    key: fullKey,
                    defaultValue: defaultValue,
                    globalValue: value,
                    workspaceValue: value,
                    workspaceFolderValue: undefined
                };
            }
        };
        return config as any;
    }) as any;
    
    console.log('✅ Initialized test configuration system with centralized VS Code config mocking');
}

/**
 * Clean up test configuration and restore original VS Code config system
 */
export function cleanupTestConfig(): void {
    setTestConfig(null);
    testConfigStore.clear();
    
    // Restore original VS Code configuration function
    if (originalGetConfiguration) {
        vscode.workspace.getConfiguration = originalGetConfiguration;
    }
    
    console.log('✅ Cleaned up test configuration system and restored VS Code config');
}

/**
 * Update test configuration
 */
export async function testConfigUpdate(key: string, value: any): Promise<void> {
    testConfigStore.set(key, value);
    console.log(`✅ Test config update: ${key} = ${JSON.stringify(value)}`);
}

/**
 * Get test configuration value
 */
export function getTestConfig<T>(key: string, defaultValue?: T): T | undefined {
    return testConfigStore.get(key) ?? defaultValue;
}

/**
 * Mock workspace configuration (deprecated - use initTestConfig instead)
 * Kept for backward compatibility
 */
export function mockWorkspaceConfiguration(): () => void {
    initTestConfig();
    
    // Return cleanup function
    return () => {
        cleanupTestConfig();
    };
}

/**
 * Test command execution helper
 */
export async function testCommandExecution(commandId: string, ...args: any[]): Promise<any> {
    try {
        const result = await vscode.commands.executeCommand(commandId, ...args);
        console.log(`✅ Command executed successfully: ${commandId}`);
        return result;
    } catch (error) {
        console.log(`⚠️ Command execution failed: ${commandId} - ${error}`);
        throw error;
    }
}

/**
 * Create mock configuration (deprecated - use initTestConfig instead)
 */
export function createMockConfig(defaults?: Record<string, any>) {
    if (defaults) {
        Object.entries(defaults).forEach(([key, value]) => {
            testConfigStore.set(key, value);
        });
    }
    return testConfigStore;
}
