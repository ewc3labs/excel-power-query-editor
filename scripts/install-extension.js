#!/usr/bin/env node

/**
 * Cross-platform extension installer script
 * Reads package.json to get name and version, then installs the VSIX
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// Read package.json to get name and version
const packageJsonPath = path.join(__dirname, '..', 'package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

const extensionName = packageJson.name;
const extensionVersion = packageJson.version;
const vsixFileName = `${extensionName}-${extensionVersion}.vsix`;

// Check if --force flag is passed
const forceFlag = process.argv.includes('--force') ? ' --force' : '';

// Construct the command
const command = `code --install-extension ${vsixFileName}${forceFlag}`;

console.log(`Installing extension: ${vsixFileName}`);
console.log(`Command: ${command}`);

try {
    // Execute the command
    execSync(command, { 
        stdio: 'inherit',
        cwd: path.join(__dirname, '..')
    });
    console.log(`✅ Successfully installed ${vsixFileName}`);
} catch (error) {
    console.error(`❌ Failed to install extension: ${error.message}`);
    process.exit(1);
}
