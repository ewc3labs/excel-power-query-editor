# Change Log

All notable changes to the "excel-power-query-editor" extension will be documented in this file.

## [0.4.3] - 2025-06-20

### Added
- **📦 VS Code Marketplace**: Published extension to VS Code Marketplace (ewc3labs.excel-power-query-editor)
- **🚀 Installation Instructions**: Updated README and USER_GUIDE with marketplace installation steps
- **⚡ Quick Start**: Added Quick Start section to README for immediate user value

### Improved
- **🎨 Extension Icon**: Optimized extension logo for better marketplace presentation
- **📚 Documentation**: Updated installation instructions to prioritize marketplace over VSIX files
- **🧹 Repository Cleanup**: Removed test folder and test files from public repository

## [0.4.2] - 2025-06-20

### Added
- **💖 Support Links**: Added "Buy Me a Coffee" support links in README, USER_GUIDE, and dedicated SUPPORT.md
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

## [0.3.1] - 2025-06-18

### Added
- **Comprehensive settings**: All configuration options implemented in package.json
- **Auto-watch settings**: `watchAlways`, `watchOffOnDelete`, `syncDeleteTurnsWatchOff`
- **Status bar integration**: Shows watch count when files are being monitored

### Fixed
- **Auto-watch behavior**: Improved reliability and user control over automatic watching

## [0.3.0] - 2025-06-17

### Added
- **File watching**: Auto-sync .m files to Excel when changes detected
- **Toggle watch**: Smart toggle command to start/stop watching files
- **Status indicators**: Visual feedback for watch status in status bar

## [0.2.2] - 2025-06-16

### Fixed
- **Sync reliability**: Improved binary blob handling and XML reconstruction
- **Comment preservation**: Ensures comments in M code are maintained during sync
- **Error handling**: Better handling of sync failures with debug information

## [0.2.1] - 2025-06-15

### Fixed
- **UTF-16 LE BOM decoding**: Proper handling of Excel DataMashup XML encoding
- **Sync accuracy**: Improved Excel file modification process

## [0.2.0] - 2025-06-14

### Added
- **Sync functionality**: Sync modified .m files back to Excel
- **Debug features**: Raw extraction and verbose logging
- **Backup system**: Automatic backups before sync operations

## [0.1.3] - 2025-06-13

### Added
- **Initial release**: Extract Power Query from Excel files to .m files
- **File naming convention**: Uses full Excel filename (e.g., `file.xlsx_PowerQuery.m`)
- **Multiple format support**: Works with .xlsx, .xlsm, and .xlsb files
- **Cross-platform**: No COM dependencies, works on Windows, macOS, Linux