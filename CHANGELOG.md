# Excel Power Query Editor

A modern, reliable VS Code extension for editing Power Query M code directly from Excel files.

---

# Changelog

All notable changes to the "excel-power-query-editor" extension will be documented in this file.

---


## [0.6.0] - 2026-08-15

### Added

- **Live sync: write to a workbook Excel already has open.** Editing a `.m` file and syncing no
  longer requires closing the workbook. When Excel has it open, the change is made through Excel
  itself, the file on disk is untouched, and the workbook is left with unsaved changes for you to
  review and save. Off by default — set `excel-power-query-editor.sync.liveWhenOpen` to `true`.

  Requested in [discussion #3](https://github.com/ewc3labs/excel-power-query-editor/discussions/3).

  **Live sync declines when the workbook has unsaved changes**, because the backup taken before a
  sync copies the file on disk and cannot contain edits that exist only in Excel. Governed by
  `excel-power-query-editor.sync.requireSavedWorkbook`, on by default. Turn it off for rapid
  iteration on scratch workbooks, accepting that a backup then holds the last state you saved rather
  than the state you were working in.

  **AutoSave has not been tested against this.** Cloud workbooks usually have it on, and it may
  persist a live write before you can review it — see
  [PQ-33](docs/project/slices/PQ-33_AutoSave_And_Live_Sync.md).


  **This is beta, and Excel is a big place.** It works by talking to Excel through COM, and Excel
  varies enormously between versions, update channels and managed environments. What was measured
  to work: several Excel instances at once (workbooks are found through the Running Object Table,
  so it does not matter which instance has yours), OneDrive and SharePoint (a synced workbook is
  registered under its cloud URL rather than the path you see), and Excel being busy (modal
  dialogs and recalculation fail COM calls, and are retried).

  What has **not** been tried: Protected View, workbooks opened from a web link, co-authored files
  with AutoSave, `excel.exe /x`, and environments where group policy restricts automation.

  If it declines to work, that is worth reporting. Set `log.level` to `debug`, reproduce, and send
  the line beginning `Excel file is locked; live sync ...` from the output channel — it says which
  branch was taken and why.

- **Excel symbols now register through the Power Query extension's API** rather than being written
  into your workspace. Nothing is added to `.vscode/`, no other extension's settings are edited,
  and it works with no folder open. They are re-registered automatically if the Power Query
  extension is installed or updated after this one.

### Changed

- **Settings are namespaced** — `watchAlways` → `watch.always`, `logLevel` → `log.level`, and so
  on. **Your existing settings are migrated automatically** on first run. The old names remain
  declared and deprecated for at least one release, so a user skipping a version still migrates.
- Logging uses a VS Code `LogOutputChannel`, so verbosity follows **Developer: Set Log Level…** and
  the log is persisted by VS Code with its own.
- One README, used by both the repository and the Marketplace listing.
- **The documentation was reorganized**, following [Klipper][klipper]'s model: flat files, one
  document per feature, and references that enumerate everything. New: `Overview`, `Installation`,
  `FAQ`, `Commands`, `Live_Sync`, `Watch_Mode`, `Backups`, `Excel_Symbols`, `Config_Reference` and
  `Config_Changes`. `USER_GUIDE.md` is now `User_Guide.md`; `CONFIGURATION.md` was replaced by the
  generated `Config_Reference.md`, with renamed settings moved to `Config_Changes.md`; the release
  summary was retired in favor of this file. **A bookmark to an old filename will not resolve** —
  start at [docs/Overview.md](docs/Overview.md).
- `Config_Reference.md` is now generated from `package.json`, and CI fails if the two disagree, so a
  setting cannot be added without appearing in the reference.

### Internal

- Documentation is checked in CI by [@ewc3labs/docs-tools][docs-tools], a small public toolkit split
  out of this repository: prose wrapped to render width, links verified (including their **case**,
  which Windows and macOS resolve happily and GitHub does not), and computed numbers refreshed from
  what they count. `npm run docs:fix` does the work; `npm run docs:check` is what CI runs. See
  [CONTRIBUTING](docs/CONTRIBUTING.md).
- The release pipeline is triggered by the tag rather than by CI completing. The previous one used
  `workflow_run`, which executes in the default branch context, so its `refs/tags/v*` conditions were
  unreachable and it silently released nothing after 2025-07-21.

### Fixed

- **`migrateLegacySettings` deleted settings rather than migrating them.** The previous
  implementation set every key in scope to `undefined` in both User and Workspace scope, preserving
  nothing, and re-ran on every version bump. It never shipped. If you ran a development build
  between 0.5.0 and now and lost your configuration, this was why.
- Extraction no longer fails when no folder is open.
- The release pipeline builds again. It had not succeeded since July 2025, which is why 0.5.1 and
  0.5.2 were never published.

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

[docs-tools]: https://github.com/ewc3labs/ewc3-docs-tools
[klipper]: https://github.com/Klipper3d/klipper
