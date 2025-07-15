<div align="center">

# ![Excel Power Query Editor](assets/excel-power-query-editor-logo-128x128.png) Excel Power Query Editor

## Changelog

All notable changes to the "excel-power-query-editor" extension will be documented in this file.

Check [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) for formatting standards.

---

**Built with 🧡 by** [![EWC3 Labs](assets/EWC3LabsLogo-blue-128x128.png)](https://github.com/ewc3labs) **EWC3 Labs**  
_A skunkworks of code, plastic, and canine gaseous emissions_

</div>

## [0.5.0-rc.2] - 2025-07-14

### 🚀 Major Performance & Feature Release

#### Added
- **NEW FEATURE: Excel Power Query Symbols System**
  - Complete Excel-specific IntelliSense support (Excel.CurrentWorkbook, Excel.Workbook, etc.)
  - Auto-installation with Power Query Language Server integration
  - Addresses gap in M Language extension (Power BI/Azure focused)
  - Configurable installation scope (workspace/folder/user/off)

#### Fixed  
- **CRITICAL: Auto-Save Performance Crisis**
  - Resolved VS Code auto-save + file watcher causing keystroke-level sync with large files
  - Intelligent debouncing based on Excel file size (not .m file size)
  - Large file handling: 3000ms → 8000ms debounce for files >10MB
- **Test Infrastructure Excellence**
  - All 71 tests passing across platforms
  - Eliminated test hangs from file dialogs and background processes
  - Auto-compilation for VS Code Test Explorer
  - Robust parameter validation and error handling

#### Changed
- **Configuration Best Practices**
  - ⚠️ **WARNING**: DO NOT enable VS Code auto-save + Extension auto-watch simultaneously
  - Recommended: `"files.autoSave": "off"` with extension file watching
  - Documented optimal performance configuration patterns

## [0.5.0] - 2025-07-11

### Added

- **New Configuration Options**:
  - `sync.openExcelAfterWrite`: Auto-launch Excel after sync operations
  - `sync.debounceMs`: Configurable sync delay (prevents duplicate syncs with CoPilot)
  - `watch.checkExcelWriteable`: Validate Excel file access before sync
  - `backup.maxFiles`: Replaces `maxBackups` with improved backup retention
- **New Commands**:
  - `Apply Recommended Defaults`: Sets optimal configuration for new users
  - `Cleanup Old Backups`: Manual backup management
- **Enhanced Error Handling**: Locked file detection with retry logic and clear user feedback
- **CoPilot Integration**: Intelligent debouncing and file hash deduplication prevents triple-sync issues

### Improved

- **Test Coverage**: 63 comprehensive tests with 100% pass rate across platforms
- **CI/CD Pipeline**: Cross-platform GitHub Actions with Ubuntu, Windows, macOS validation
- **Development Environment**: Complete DevContainer setup with pre-configured dependencies
- **Documentation**: Comprehensive USER_GUIDE.md, CONFIGURATION.md, and CONTRIBUTING.md

### Fixed

- **Settings System**: Centralized VS Code API mocking for reliable test environment
- **Command Registration**: All commands properly registered and available in test environment
- **Watch Mode**: Improved debouncing prevents unnecessary sync operations
- **Configuration Migration**: Automatic v0.4.x settings migration to v0.5.0 structure

### Technical

- **Quality Gates**: ESLint, TypeScript, and test validation in CI/CD
- **Cross-Platform**: Ubuntu 22.04, Windows Server 2022, macOS 14 compatibility verified
- **Artifact Management**: VSIX packaging with 30-day retention

---

## [0.4.3] - 2025-06-20

### Added

- **VS Code Marketplace**: Published extension to VS Code Marketplace (ewc3labs.excel-power-query-editor)
- **Installation Instructions**: Updated README and USER_GUIDE with marketplace installation steps
- **Quick Start**: Added Quick Start section to README for immediate user value

### Improved

- **Extension Icon**: Optimized extension logo for better marketplace presentation
- **Documentation**: Updated installation instructions to prioritize marketplace over VSIX files
- **Repository Cleanup**: Removed test folder and test files from public repository

## [0.4.2] - 2025-06-20

### Added

- **Support Links**: Added "Buy Me a Coffee" support links in README, USER_GUIDE, and dedicated SUPPORT.md
- **Extension Pack**: Automatically installs Microsoft Power Query / M Language extension (`powerquery.vscode-powerquery`)
- **Better Categories**: Changed from "Other" to "Programming Languages", "Data Science", "Formatters"
- **Keywords**: Added searchable keywords ("excel", "power query", "m language", "data analysis", "etl") for better marketplace discoverability
- **Documentation Links**: Prominently featured links to USER_GUIDE.md and CONFIGURATION.md in README
- **Package.json Metadata**: Added bugs, homepage, and sponsor URLs for better extension page experience

### Improved

- **README**: Added required extension warning, complete documentation links, and professional support section
- **USER_GUIDE**: Updated to mention required Power Query extension for proper M language support
- **Extension Recommendations**: Clear guidance on required vs optional companion extensions
- **SUPPORT.md**: Dedicated support file following GitHub conventions

## [0.4.1] - 2025-06-20

### Added

- **Auto-watch initialization**: Scans for .m files on extension activation when `watchAlways` is enabled
- **Hybrid activation**: Always activate on startup but only auto-watch if setting is enabled
- **Performance limits**: Auto-watch limited to 20 files to prevent performance issues

### Fixed

- **Activation events**: Added `"onStartupFinished"` for proper startup behavior
- **Auto-watch reliability**: Improved restoration of watch state after VS Code reload

## [0.4.0] - 2025-06-19

### Added

- **Backup management**: Configurable max backups with auto-cleanup
- **Cleanup command**: Manual "Cleanup Old Backups" command for Excel files
- **Custom backup locations**: Support for same folder, temp folder, or custom paths
- **Backup retention**: Automatically delete old backups when limit exceeded

### Improved

- **Settings organization**: Comprehensive settings for backup management
- **User experience**: Better feedback for backup and cleanup operations

## [Initial Release] - 2025-06-13

### Added

- **Core functionality**: Extract Power Query from Excel files to .m files
- **File format support**: Works with .xlsx, .xlsm, and .xlsb files
- **Sync capability**: Sync modified .m files back to Excel
- **File watching**: Auto-sync .m files to Excel when changes detected
- **Cross-platform**: No COM dependencies, works on Windows, macOS, Linux
- **Backup system**: Automatic backups before sync operations
