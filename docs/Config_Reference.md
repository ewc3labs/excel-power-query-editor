# Configuration Reference

Every setting this extension contributes, generated from `package.json` so it cannot drift from the
software. Settings are edited in VS Code's settings UI, or in `settings.json` using the full key.

Upgrading from 0.5.x? See [Config Changes](Config_Changes.md) — the settings were renamed, and your
existing values are migrated automatically.

## Sync

How, and where, your `.m` changes are written back.

### `excel-power-query-editor.sync.debounceMs`

| | |
| --- | --- |
| Type | `number` |
| Default | `500` |

Milliseconds to debounce file saves before sync. Prevents duplicate syncs in rapid succession.

### `excel-power-query-editor.sync.deleteAlwaysConfirm`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Show a confirmation dialog before syncing and deleting the .m file. Uncheck to perform without confirmation.

### `excel-power-query-editor.sync.liveWhenOpen`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |

**Beta — Windows + Excel only.** When the workbook is already open in Excel, write the query changes straight into it through Excel itself instead of failing or waiting for the file to be closed. The file on disk is not touched — the workbook is left with unsaved changes for you to review and save. Off by default. This talks to Excel through COM, and Excel varies a great deal between versions, update channels and managed environments. If it declines to work, set `excel-power-query-editor.log.level` to `debug` and [open an issue](https://github.com/ewc3labs/excel-power-query-editor/issues) with the `live sync ...` line from the output channel. Closed workbooks are unaffected, and so is extraction.

### `excel-power-query-editor.sync.openExcelAfterWrite`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |

[PLANNED FEATURE] Automatically open the Excel file after a successful sync.

### `excel-power-query-editor.sync.requireSavedWorkbook`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

**On by default. Leave it on for the safest behavior; turn it off for rapid iteration.** When the workbook is open in Excel with unsaved changes, live sync declines and asks you to save it first. The reason is that the backup taken before a sync is a copy of the file **on disk** — the last saved state. Edits sitting unsaved in Excel are in no file anywhere, so writing over them destroys the only copy while the backup looks like protection. **Turning this off is a real trade, not a formality.** It is genuinely useful when you are beating up prototype queries in a scratch workbook and do not want to save between every sync. What you give up is that a failed or unwanted sync can no longer be undone from a backup — the backup will contain the last state you saved, not the state you were working in. Only applies to live sync. Closed workbooks are written to disk with a backup as usual. **In practice this protects LOCAL workbooks.** A workbook with AutoSave on - which most OneDrive and SharePoint files have - is saved by Excel within seconds, so it is almost never sitting dirty for this to catch. That is not a gap: AutoSave has already put those edits on disk, where the backup can capture them. The unsaved-work danger is a local-file problem, and a cloud workbook with AutoSave switched off is protected here exactly like a local one.

### `excel-power-query-editor.sync.timeout`

| | |
| --- | --- |
| Type | `number` |
| Default | `30000` |

Time in milliseconds before a sync attempt is aborted.

## Watch

Syncing automatically when a file is saved.

### `excel-power-query-editor.watch.always`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |

Automatically start watching when extracting Power Query files

### `excel-power-query-editor.watch.checkExcelWriteable`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Before syncing, check if Excel file is writable. Warn or retry if locked.

### `excel-power-query-editor.watch.maxFiles`

| | |
| --- | --- |
| Type | `number` |
| Default | `25` |

Maximum number of .m files to auto-watch when watchAlways is enabled. Prevents performance issues with large workspaces.

### `excel-power-query-editor.watch.offOnDelete`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Stop watching a .m file if it is deleted from disk.

## Backups

What is kept before a workbook is written, and for how long.

### `excel-power-query-editor.backup.autoBackupBeforeSync`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Automatically create a backup of the Excel file before syncing from .m.

### `excel-power-query-editor.backup.autoCleanup`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Enable automatic deletion of old backups when the number exceeds maxBackups.

### `excel-power-query-editor.backup.customPath`

| | |
| --- | --- |
| Type | `string` |
| Default | `""` |

Path to use if backupLocation is set to "custom". Can be relative to the workspace root.

### `excel-power-query-editor.backup.location`

| | |
| --- | --- |
| Type | `string` |
| Default | `"sameFolder"` |
| Values | `sameFolder` · `tempFolder` · `custom` |

Folder to store backup files: same as Excel file, system temp folder, or a custom path.

### `excel-power-query-editor.backup.maxFiles`

| | |
| --- | --- |
| Type | `number` |
| Default | `5` |

Maximum number of backup files to retain per Excel file. Older backups are deleted when exceeded.

## Logging

What the extension records, and where.

### `excel-power-query-editor.log.level`

| | |
| --- | --- |
| Type | `string` |
| Default | `"info"` |
| Values | `none` · `error` · `warn` · `info` · `verbose` · `debug` |

Set the logging level for the Excel Power Query Editor extension. Replaces legacy verboseMode and debugMode settings.

### `excel-power-query-editor.log.showStatusBarInfo`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |

Display sync and watch status indicators in the VS Code status bar.

## Internal

Written by the extension. Not intended to be set by hand.

### `excel-power-query-editor.xtn.level`

| | |
| --- | --- |
| Type | `string` |
| Default | `""` |

Internal migration marker for extension settings. Do not edit.

## Deprecated

Renamed in 0.6.0. Still read, still migrated, and removed after a release or two. See [Config Changes](Config_Changes.md).

### `excel-power-query-editor.autoBackupBeforeSync`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |
| Status | **Deprecated** |

Automatically create a backup of the Excel file before syncing from .m.

> Deprecated: use excel-power-query-editor.backup.autoBackupBeforeSync. Migrated automatically.

### `excel-power-query-editor.autoCleanupBackups`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |
| Status | **Deprecated** |

Enable automatic deletion of old backups when the number exceeds maxBackups.

> Deprecated: use excel-power-query-editor.backup.autoCleanup. Migrated automatically.

### `excel-power-query-editor.backupLocation`

| | |
| --- | --- |
| Type | `string` |
| Default | `"sameFolder"` |
| Values | `sameFolder` · `tempFolder` · `custom` |
| Status | **Deprecated** |

Folder to store backup files: same as Excel file, system temp folder, or a custom path.

> Deprecated: use excel-power-query-editor.backup.location. Migrated automatically.

### `excel-power-query-editor.customBackupPath`

| | |
| --- | --- |
| Type | `string` |
| Default | `""` |
| Status | **Deprecated** |

Path to use if backupLocation is set to "custom". Can be relative to the workspace root.

> Deprecated: use excel-power-query-editor.backup.customPath. Migrated automatically.

### `excel-power-query-editor.debugMode`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |
| Status | **Deprecated** |

[DEPRECATED] Use logLevel instead. Enable debug-level logging and write internal debug files to disk.

> Deprecated: use excel-power-query-editor.log.level set to 'debug'. Migrated automatically.

### `excel-power-query-editor.logLevel`

| | |
| --- | --- |
| Type | `string` |
| Default | `"info"` |
| Values | `none` · `error` · `warn` · `info` · `verbose` · `debug` |
| Status | **Deprecated** |

Set the logging level for the Excel Power Query Editor extension. Replaces legacy verboseMode and debugMode settings.

> Deprecated: use excel-power-query-editor.log.level. Migrated automatically.

### `excel-power-query-editor.showStatusBarInfo`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |
| Status | **Deprecated** |

Display sync and watch status indicators in the VS Code status bar.

> Deprecated: use excel-power-query-editor.log.showStatusBarInfo. Migrated automatically.

### `excel-power-query-editor.syncDeleteAlwaysConfirm`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |
| Status | **Deprecated** |

Show a confirmation dialog before syncing and deleting the .m file. Uncheck to perform without confirmation.

> Deprecated: use excel-power-query-editor.sync.deleteAlwaysConfirm. Migrated automatically.

### `excel-power-query-editor.syncTimeout`

| | |
| --- | --- |
| Type | `number` |
| Default | `30000` |
| Status | **Deprecated** |

Time in milliseconds before a sync attempt is aborted.

> Deprecated: use excel-power-query-editor.sync.timeout. Migrated automatically.

### `excel-power-query-editor.verboseMode`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |
| Status | **Deprecated** |

[DEPRECATED] Use logLevel instead. Output detailed logs to the VS Code Output panel (recommended for troubleshooting).

> Deprecated: use excel-power-query-editor.log.level set to 'verbose'. Migrated automatically.

### `excel-power-query-editor.watchAlways`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `false` |
| Status | **Deprecated** |

Automatically start watching when extracting Power Query files

> Deprecated: use excel-power-query-editor.watch.always. Migrated automatically.

### `excel-power-query-editor.watchAlwaysMaxFiles`

| | |
| --- | --- |
| Type | `number` |
| Default | `25` |
| Status | **Deprecated** |

Maximum number of .m files to auto-watch when watchAlways is enabled. Prevents performance issues with large workspaces.

> Deprecated: use excel-power-query-editor.watch.maxFiles. Migrated automatically.

### `excel-power-query-editor.watchOffOnDelete`

| | |
| --- | --- |
| Type | `boolean` |
| Default | `true` |
| Status | **Deprecated** |

Stop watching a .m file if it is deleted from disk.

> Deprecated: use excel-power-query-editor.watch.offOnDelete. Migrated automatically.

---

_Generated by `scripts/generate-config-reference.js`. Do not edit by hand — change `package.json` and re-run `npm run docs:config`._
