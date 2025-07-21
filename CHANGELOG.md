# Excel Power Query Editor

A modern, reliable VS Code extension for editing Power Query M code directly from Excel files.

---

# Changelog

All notable changes to the "excel-power-query-editor" extension will be documented in this file.

---


## [0.5.0] - 2025-07-20

### 🎯 Marketplace Release - Professional Logging, Auto-Watch Enhancements, Symbols, and Legacy Settings Migration

#### Added
- **Excel Power Query Symbols System**
  - Complete Excel-specific IntelliSense support (Excel.CurrentWorkbook, Excel.Workbook, etc.)
  - Auto-installation with Power Query Language Server integration
  - Configurable installation scope (workspace/folder/user/off)
- **Professional Logging System**
  - Emoji-enhanced logging with visual level indicators (🪲🔍ℹ️✅⚠️❌)
  - Six configurable log levels: none, error, warn, info, verbose, debug
  - Automatic emoji support detection for VS Code environments
  - Context-aware logging with function-specific prefixes
  - Environment detection and settings dump for debugging
- **Intelligent Auto-Watch System**
  - Configurable auto-watch file limits (`watchAlways.maxFiles`: 1-100, default 25)
  - Prevents performance issues in large workspaces with many .m files
  - Smart file discovery with Excel file matching validation
  - Detailed logging of skipped files and initialization progress
- **Enhanced Excel Symbols Integration**
  - Three-step Power Query settings update for immediate effect
  - Delete/pause/reset sequence forces Language Server reload
  - Ensures new symbols take effect without VS Code restart
  - Cross-platform directory path handling
- **Legacy Settings Migration**
  - Automatic migration of deprecated settings (`debugMode`, `verboseMode`) to new `logLevel` with user notification
- **New Commands**
  - `Apply Recommended Defaults`: Sets optimal configuration for new users
  - `Cleanup Old Backups`: Manual backup management

#### Fixed & Improved
- **Auto-Save Performance**
  - Resolved VS Code auto-save + file watcher causing keystroke-level sync with large files
  - Intelligent debouncing based on Excel file size (not .m file size)
  - Large file handling: 3000ms → 8000ms debounce for files >10MB
- **Test Infrastructure**
  - 74 comprehensive tests with 100% pass rate, including legacy settings migration
  - Eliminated test hangs from file dialogs and background processes
  - Auto-compilation for VS Code Test Explorer
  - Robust parameter validation and error handling
- **Configuration System**
  - Fixed `watchAlwaysMaxFiles` setting validation (was incorrectly named `watchAlways.maxFiles`)
  - VS Code settings now properly accept numeric input for auto-watch file limits
  - Resolved "Value must be a number" error in extension settings
  - v0.4.x settings (`debugMode`, `verboseMode`) are now automatically migrated to the new `logLevel` system
- **Logging System Consistency**
  - Fixed context naming inconsistencies (ExtractFromExcel → extractFromExcel)
  - Replaced generic contexts with specific function names
  - Optimized log levels for better user experience
  - Eliminated double logging patterns
- **Auto-Watch Performance**
  - Intelligent file limit enforcement prevents extension overwhelm
  - Better handling of workspaces with many test fixtures
  - Improved startup time with configurable limits
- **Settings System**
  - Centralized VS Code API mocking for reliable test environment
  - All commands properly registered and available in test environment
  - Improved debouncing prevents unnecessary sync operations
  - Automatic v0.4.x settings migration to v0.5.0 structure

#### Changed & Technical
- **VS Code Marketplace Ready**
  - Professional user experience with polished logging
  - Enhanced settings documentation
  - Optimal default configurations for production use
- **Test Coverage**
  - 74 comprehensive tests with 100% pass rate, including legacy settings migration
- **CI/CD Pipeline**
  - Cross-platform GitHub Actions with Ubuntu, Windows, macOS validation
- **Development Environment**
  - Complete DevContainer setup with pre-configured dependencies
- **Documentation**
  - Comprehensive USER_GUIDE.md, CONFIGURATION.md, and CONTRIBUTING.md
- **Quality Gates**
  - ESLint, TypeScript, and test validation in CI/CD
- **Cross-Platform**
  - Ubuntu 22.04, Windows Server 2022, macOS 14 compatibility verified
- **Artifact Management**
  - VSIX packaging with 30-day retention

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
