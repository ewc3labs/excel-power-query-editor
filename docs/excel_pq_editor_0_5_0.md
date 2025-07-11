## Excel Power Query Editor v0.5.0 Release Plan

### 🚀 Major Goals (Narrative)

v0.5.0 is the first polish-and-harden release now that the extension has crossed 100 installs. It focuses on eliminating core usability issues, improving onboarding docs, fixing test harness limitations, and preparing the repo for public contributions and continued growth. The goal is to make it frictionless for first-time users and robust for day-to-day use in version-controlled Excel Power Query development.

---

### ✅ Critical Bugs (Must-Fix Before Release)

#### 🕱️ Right-click handler not registering on sidebar

- VS Code API does not trigger context menu commands when `.m` file is clicked in Explorer unless editor has focus
- 🛠️ Fix: Adjust `when` clauses and activation logic so file tree clicks also initialize command targets

#### ⚙️ `settings.json` ignored in test harness

- Tests currently can't load user/workspace settings, breaking settings-dependent features
- 🛠️ Fix: Inject mocked `vscode.workspace.getConfiguration()` for test environment
- 🔹 Move `test` folder to repo root (was under `src/`) to align with standard layout

#### ♻️ CoPilot Agent mode causes triple sync

- Save → Sync (expected)
- Agent diff → Sync
- Accept diff → Sync again
- 🛠️ Fix: Add debounce, or dedupe within {configurable}ms using file hash/timestamp (don't sync if hash matches)

#### 📄 Silent failure on open Excel file

- If Excel locks the file, sync fails with no UI warning
- 🛠️ Fix: Detect locked file via try/catch or fs.open with error codes; show warning or retry
- 🛠️ Fix: Configurable watch Excel file for availability, write after debounce delay.

---

### 📋 Docs Overhaul Tasks

| Section          | Status | Fix                                                           |
| ---------------- | ------ | ------------------------------------------------------------- |
| Docs Structure   | ✅     | Add `docs/` folder, move non-README documentation inside      |
| README           | 🔄     | Split Features vs Usage vs Config vs Install                  |
| Usage Guide      | ❌     | Show typical `.m` file lifecycle                              |
| Watch Mode       | ❌     | Explain backup rotation, timestamp caps, auto-enable setting  |
| Right-Click Sync | ❌     | Emphasize you must click _inside the editor_                  |
| Settings         | ❌     | Add table of options, defaults, scopes                        |
| GIF/MP4          | ❌     | Add short demo (show extract → edit → save → auto-sync)       |
| CLI Install      | ❌     | `code --install-extension ewc3labs.excel-power-query-editor`  |
| Known Issues     | ❌     | Can't sync to open-in-Excel file (implement watch Excel file) |

---

### 🔧 New Features

- `sync.openExcelAfterWrite`: If enabled, launches Excel after sync
- `watch.backup.maxFiles`: Sets retention cap for auto-backups
- `sync.debounceMs`: Configurable debounce delay before sync
- `watch.checkExcelWriteable`: If true, re-checks Excel file write access before sync
- **Apply Recommended Defaults command**: New command in Command Palette to set smart recommended defaults after upgrade or first install

> New settings are introduced in 0.5.0 with sensible defaults. Users who installed before 0.5.0 may wish to run the new "Apply Recommended Defaults" command from the Command Palette to update their settings.

---

### ⚙️ Configuration Settings (Proposed Revisions)

| Setting Key                                          | Type      | Default      | Current Description                                                                             | Proposed Description                                                                                           | Notes                                              |
| ---------------------------------------------------- | --------- | ------------ | ----------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- | -------------------------------------------------- |
| `excel-power-query-editor.watchAlways`               | `boolean` | `false`      | Automatically start watching when extracting Power Query files                                  | Automatically enable watch mode after extracting Power Query files.                                            | OK                                                 |
| `excel-power-query-editor.watchOffOnDelete`          | `boolean` | `true`       | Automatically stop watching when the .m file is deleted                                         | Stop watching a `.m` file if it is deleted from disk.                                                          | OK                                                 |
| `excel-power-query-editor.syncDeleteTurnsWatchOff`   | `boolean` | `true`       | Stop watching when using 'Sync & Delete'                                                        | Automatically disable watch mode after using **Sync and Delete**.                                              | Redundant with `watchOffOnDelete`; remove in 0.5.0 |
| `excel-power-query-editor.syncDeleteAlwaysConfirm`   | `boolean` | `true`       | Always ask for confirmation before 'Sync & Delete' (uncheck to skip confirmation)               | Show a confirmation dialog before syncing and deleting the `.m` file. Uncheck to perform without confirmation. | OK                                                 |
| `excel-power-query-editor.verboseMode`               | `boolean` | `false`      | Show detailed output in the Output panel                                                        | Output detailed logs to the VS Code Output panel (recommended for troubleshooting).                            | OK                                                 |
| `excel-power-query-editor.autoBackupBeforeSync`      | `boolean` | `true`       | Create automatic backups before syncing to Excel                                                | Automatically create a backup of the Excel file before syncing from `.m`.                                      | OK                                                 |
| `excel-power-query-editor.backupLocation`            | `enum`    | `sameFolder` | Where to store backup files                                                                     | Folder to store backup files: same as Excel file, system temp folder, or a custom path.                        | OK                                                 |
| `excel-power-query-editor.customBackupPath`          | `string`  | `""`         | Custom path for backups (when backupLocation is 'custom')                                       | Path to use if `backupLocation` is set to `"custom"`. Can be relative to the workspace root.                   | OK                                                 |
| `excel-power-query-editor.maxBackups`                | `number`  | `5`          | Maximum number of backup files to keep per Excel file (older backups are automatically deleted) | Maximum number of backup files to retain per Excel file. Older backups are deleted when exceeded.              | Rename to `backup.maxFiles` in 0.5.0               |
| `excel-power-query-editor.autoCleanupBackups`        | `boolean` | `true`       | Automatically delete old backup files when exceeding maxBackups limit                           | Enable automatic deletion of old backups when the number exceeds `maxBackups`.                                 | OK                                                 |
| `excel-power-query-editor.syncTimeout`               | `number`  | `30000`      | Timeout in milliseconds for sync operations                                                     | Time in milliseconds before a sync attempt is aborted.                                                         | OK                                                 |
| `excel-power-query-editor.debugMode`                 | `boolean` | `false`      | Enable debug logging and save debug files                                                       | Enable debug-level logging and write internal debug files to disk.                                             | OK                                                 |
| `excel-power-query-editor.showStatusBarInfo`         | `boolean` | `true`       | Show watch status and sync info in the status bar                                               | Display sync and watch status indicators in the VS Code status bar.                                            | OK                                                 |
| `excel-power-query-editor.sync.openExcelAfterWrite`  | `boolean` | `false`      | _(New setting)_                                                                                 | Automatically open the Excel file after a successful sync.                                                     | New setting                                        |
| `excel-power-query-editor.sync.debounceMs`           | `number`  | `500`        | _(New setting)_                                                                                 | Milliseconds to debounce file saves before sync. Prevents duplicate syncs in rapid succession.                 | New setting                                        |
| `excel-power-query-editor.watch.checkExcelWriteable` | `boolean` | `true`       | _(New setting)_                                                                                 | Before syncing, check if Excel file is writable. Warn or retry if locked.                                      | New setting                                        |

---

### 🔪 Dev / Test Improvements

#### Devcontainer ✅ COMPLETED

- ✅ Node 22, VS Code, this extension dev environment ready
- ✅ Power Query syntax extension auto-installed
- ✅ Dev container with all required dependencies preloaded
- ✅ VS Code Tasks for test, lint, build, package extension
- 🔄 **TODO**: Add `.xlsx`, `.xlsm`, `.xlsb` sample files for fixture tests

#### Tests 🚧 IN PROGRESS

- 🔄 **CURRENT**: Move test folder from `src/test/` to `/test` root ✅ DONE
- 🔄 **CURRENT**: Create test fixtures with Excel files (with and without PQ)
- ❌ Mock `settings.json` using injected config
- ❌ Validate watch mode with file change + sync
- ❌ Trigger extract → sync flow across formats
- ❌ Add test for file locked scenario
- ❌ Add recommended defaults test validation

#### GitHub Actions

- Lint / compile
- Run headless watch mode test
- Run sync test suite

---

### 💬 Community + Marketplace

- Revise Marketplace tags: `Excel`, `Power Query`, `CoPilot`, `Data Engineering`
- README badges: install count, last published, open issues
- Discussions and issue templates
- Add `docs/` folder for usage, settings, and architecture docs

---

### 📦 Internal Project Tasks

- ✅ Add Docker dev container with all required dependencies preloaded
- ✅ Add VS Code Tasks for test, lint, build, extract/sync fixture files
- ✅ Move documentation to `docs/` folder structure
- 🔄 **NEXT**: Create test fixtures (Excel files with/without Power Query)
- ❌ Add `Apply Recommended Settings` command to initialize smart defaults on first run
