#!/usr/bin/env node

/**
 * Smart version bumping script for Excel Power Query Editor
 * Handles semantic versioning based on git context and commit messages
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function getCurrentVersion() {
  const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
  return packageJson.version;
}

function updatePackageVersion(newVersion) {
  const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
  packageJson.version = newVersion;
  fs.writeFileSync('package.json', JSON.stringify(packageJson, null, 2) + '\n');
  console.log(`✅ Updated package.json version to ${newVersion}`);
}

function getCommitsSinceLastTag() {
  try {
    const lastTag = execSync('git describe --tags --abbrev=0', { encoding: 'utf8' }).trim();
    const commits = execSync(`git log ${lastTag}..HEAD --oneline`, { encoding: 'utf8' });
    return commits.split('\n').filter(line => line.trim());
  } catch (error) {
    // No tags yet, get all commits
    const commits = execSync('git log --oneline', { encoding: 'utf8' });
    return commits.split('\n').filter(line => line.trim());
  }
}

function determineVersionBump(commits) {
  let hasMajor = false;
  let hasMinor = false;
  let hasPatch = false;

  for (const commit of commits) {
    const message = commit.toLowerCase();
    
    // Breaking changes (MAJOR)
    if (message.includes('breaking') || message.includes('!:') || message.includes('major:')) {
      hasMajor = true;
    }
    // New features (MINOR)
    else if (message.includes('feat:') || message.includes('feature:') || message.includes('add:')) {
      hasMinor = true;
    }
    // Bug fixes and other changes (PATCH)
    else if (message.includes('fix:') || message.includes('patch:') || message.includes('update:')) {
      hasPatch = true;
    }
  }

  if (hasMajor) {
    return 'major';
  }
  if (hasMinor) {
    return 'minor';
  }
  if (hasPatch) {
    return 'patch';
  }
  return 'patch'; // Default to patch for any changes
}

function bumpVersion(version, type) {
  const [major, minor, patch] = version.split('.').map(Number);
  
  switch (type) {
    case 'major':
      return `${major + 1}.0.0`;
    case 'minor':
      return `${major}.${minor + 1}.0`;
    case 'patch':
      return `${major}.${minor}.${patch + 1}`;
    default:
      return version;
  }
}

function main() {
  const args = process.argv.slice(2);
  const currentVersion = getCurrentVersion();
  
  console.log(`📦 Current version: ${currentVersion}`);

  // If version is explicitly provided, use it
  if (args.length > 0) {
    const newVersion = args[0];
    updatePackageVersion(newVersion);
    console.log(`🎯 Set version to: ${newVersion}`);
    return;
  }

  // Auto-determine version bump based on commits
  try {
    const commits = getCommitsSinceLastTag();
    console.log(`📝 Found ${commits.length} commits since last tag`);
    
    if (commits.length === 0) {
      console.log('ℹ️ No new commits found, keeping current version');
      return;
    }

    const bumpType = determineVersionBump(commits);
    const newVersion = bumpVersion(currentVersion, bumpType);
    
    console.log(`🔄 Determined bump type: ${bumpType}`);
    console.log(`📈 New version: ${newVersion}`);
    
    updatePackageVersion(newVersion);
    
    // Show sample commits that influenced the decision
    console.log('\n📋 Recent commits analyzed:');
    commits.slice(0, 5).forEach(commit => {
      console.log(`  • ${commit}`);
    });
    
  } catch (error) {
    console.error('❌ Error analyzing commits:', error.message);
    console.log('ℹ️ Falling back to patch version bump');
    
    const newVersion = bumpVersion(currentVersion, 'patch');
    updatePackageVersion(newVersion);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  getCurrentVersion,
  updatePackageVersion,
  bumpVersion,
  determineVersionBump
};
