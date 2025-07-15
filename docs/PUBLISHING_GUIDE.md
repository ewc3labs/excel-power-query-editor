# VS Code Marketplace Publishing Guide

This guide covers the complete process for publishing the Excel Power Query Editor extension to the VS Code Marketplace using **GitHub Actions automation**.

## 🚀 Automated Publishing (Recommended)

The project includes a comprehensive GitHub Actions workflow that automates the entire release process. Here's how to set it up and use it:

### 1. Setup GitHub Secrets

#### Create Visual Studio Marketplace Personal Access Token (PAT):

1. **Go to Visual Studio Marketplace**: https://marketplace.visualstudio.com/manage
2. **Sign in** with your Microsoft account (same account used for publishing)
3. **Access Personal Access Tokens**:
   - Click your profile/publisher name in the top right
   - Select "Personal Access Tokens"
   - Or go directly to: https://marketplace.visualstudio.com/manage/publishers/ewc3labs

4. **Create New Token**:
   - Click "New Token" or "Create Token"
   - **Name**: `VS Code Extension Publishing - Excel Power Query Editor`
   - **Organization**: Select your organization or "All accessible organizations"
   - **Expiration**: Choose duration (recommended: 1 year)
   - **Scopes**: Select "Marketplace"

5. **Required Scopes**:
   - ✅ **Marketplace**: `Publish` (this gives extension publish/update permissions)

6. **Copy the Token**: 
   - **CRITICAL**: Copy the token immediately - you cannot view it again!

#### Add Token as GitHub Secret:

✅ **Already Configured**: Organization-level `VSCE_PAT` secret is set up and ready!

The `VSCE_PAT` token has been configured at the **ewc3labs organization level**, which means:
- 🔄 **Automatic Access**: All repositories in the organization inherit this secret
- 🛡️ **Centralized Management**: Update once, applies everywhere
- 🚀 **Ready to Use**: No additional configuration needed

*For reference, if you need to update or recreate the secret:*
1. **Go to GitHub Organization**: https://github.com/orgs/ewc3labs/settings/secrets/actions
2. **Edit the existing `VSCE_PAT` secret** or create a new one
3. **All repos automatically inherit** organization secrets

### 2. Marketplace Publishing Status

✅ **Marketplace Publishing is ENABLED and Ready!**

The GitHub Actions workflow is fully configured with:
- ✅ **Organization-level VSCE_PAT secret** configured
- ✅ **Automatic marketplace publishing** for tagged releases
- ✅ **Conditional publishing logic** (only stable releases, not pre-releases)
- ✅ **Error handling** with helpful feedback

**No additional setup required** - you can create a release tag and the workflow will automatically publish to the Visual Studio Marketplace!

### 3. Release Process

#### For Pre-releases (Testing):

1. **Create Release Branch**:
   ```bash
   git checkout -b release/v0.5.0
   git push origin release/v0.5.0
   ```

2. **GitHub Actions will automatically**:
   - ✅ Run all tests
   - ✅ Build and package the extension
   - ✅ Create a pre-release with version `0.5.0-rc.N`
   - ✅ Upload VSIX file
   - ⏭️ Skip marketplace publishing (pre-release only)

#### For Final Releases:

1. **Create and Push Tag**:
   ```bash
   git tag v0.5.0
   git push origin v0.5.0
   ```

2. **GitHub Actions will automatically**:
   - ✅ Run all tests
   - ✅ Build and package the extension
   - ✅ Publish to VS Code Marketplace
   - ✅ Create GitHub Release
   - ✅ Upload VSIX file
   - ✅ Generate changelog

#### Manual Release Trigger:

You can also trigger releases manually from GitHub:

1. **Go to Actions tab** in GitHub
2. **Select "🚀 Release Pipeline"**
3. **Click "Run workflow"**
4. **Choose release type**: prerelease, release, or hotfix

### 4. Current Release Workflow Features

The existing workflow (`release.yml`) includes:

- **🔍 Smart Release Detection**: Automatically determines release type based on branch/tag
- **🏗️ Multi-platform Testing**: Tests on Ubuntu (can extend to Windows/macOS)
- **📦 Dynamic Versioning**: Handles pre-releases, RCs, and final versions
- **🚀 Conditional Publishing**: Only publishes stable releases to marketplace
- **📋 Automatic Changelogs**: Generates release notes from git commits
- **🎯 Release Summary**: Provides detailed pipeline results

### 5. Version Strategy

| Branch/Tag | Version Format | Marketplace | GitHub Release |
|------------|----------------|-------------|----------------|
| `release/v0.5.0` | `0.5.0-rc.N` | ❌ No | ✅ Pre-release |
| `v0.5.0` tag | `0.5.0` | ✅ Yes | ✅ Release |
| `main` branch | `0.5.0-dev.N` | ❌ No | ❌ No |

## 🔧 Manual Publishing (Backup Method)

If you need to publish manually (for testing or emergency releases):

### Prerequisites

#### Install vsce:
```bash
npm install -g @vscode/vsce
```

#### Login with PAT:
```bash
vsce login ewc3labs
# Enter your Personal Access Token when prompted
```

### Manual Publishing Steps:

```bash
# 1. Update version
npm version 0.5.0 --no-git-tag-version

# 2. Test build
npm test
npm run compile
vsce package

# 3. Publish
vsce publish
```

## 📋 Pre-Release Checklist

Before triggering any release:

- [ ] All tests passing: `npm test`
- [ ] Code compiles cleanly: `npm run compile`
- [ ] No linting errors: `npm run lint`
- [ ] CHANGELOG.md updated with release notes
- [ ] README.md reflects latest features
- [ ] Version number updated in package.json
- [x] ✅ GitHub secrets configured (organization-level VSCE_PAT)
- [x] ✅ Marketplace publishing enabled in workflow

## 🚨 Emergency Release Process

For critical hotfixes:

1. **Create hotfix branch**:
   ```bash
   git checkout -b hotfix/v0.5.1
   # Make critical fixes
   git commit -m "fix: critical bug fix"
   git push origin hotfix/v0.5.1
   ```

2. **Use manual workflow dispatch**:
   - Go to GitHub Actions
   - Select "🚀 Release Pipeline"
   - Run workflow with "hotfix" option

3. **Tag when ready**:
   ```bash
   git tag v0.5.1
   git push origin v0.5.1
   ```

## 🎯 Quick Action Items for v0.5.0 Release

✅ **All Setup Complete - Ready to Release!**

1. **✅ VSCE_PAT secret configured** (organization-level)
2. **✅ Marketplace publishing enabled** in release.yml
3. **✅ Documentation updated** (this guide and all others)
4. **✅ All features implemented** and tested
5. **🚀 Ready to release**: `git tag v0.5.0 && git push origin v0.5.0`

**Just run the tag command above to trigger automated publishing!** 🎉

---

## 🔗 Quick Reference

**Publisher**: `ewc3labs`  
**Extension ID**: `ewc3labs.excel-power-query-editor`  
**Marketplace URL**: https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor  
**Management URL**: https://marketplace.visualstudio.com/manage/publishers/ewc3labs  
**GitHub Releases**: https://github.com/ewc3labs/excel-power-query-editor/releases

**Key Commands**:
```bash
# Manual release
git tag v0.5.0 && git push origin v0.5.0

# Pre-release testing
git checkout -b release/v0.5.0 && git push origin release/v0.5.0

# Check workflow status
gh workflow list
gh run list --workflow=release.yml
```
