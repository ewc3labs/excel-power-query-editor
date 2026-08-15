# Commands

Everything the extension adds to the Command Palette and the context menus.

Open the palette with `Ctrl+Shift+P` (`Cmd+Shift+P` on macOS) and type "Excel Power Query". Most of
these are also on the right-click menu for the file type they apply to, which is usually the faster
route.

## Working with queries

### Extract Power Query from Excel

`excel-power-query-editor.extractFromExcel` — right-click an `.xlsx`, `.xlsm` or `.xlsb`

Reads the workbook and writes a `.m` file beside it containing every query it holds. Excel does not
need to be installed, and the workbook may be open at the time — reading is never blocked, only
writing.

Select several workbooks and it extracts from all of them.

### Sync Power Query to Excel

`excel-power-query-editor.syncToExcel` — right-click a `.m` file

Writes your `.m` back to the workbook it came from. A backup is taken first (see
[Backups](Backups.md)).

Two routes, chosen per file:

- **The workbook is closed** — the query data inside the file is rewritten directly.
- **Excel has it open** and [Live Sync](Live_Sync.md) is enabled — the change is made through Excel,
  and the file on disk is untouched.

### Sync & Delete

`excel-power-query-editor.syncAndDelete`

Syncs, then deletes the `.m` file. For finishing with a query you extracted for one edit. Asks
first, unless `sync.deleteAlwaysConfirm` is off.

## Watching

See [Watch Mode](Watch_Mode.md).

### Watch File for Changes

`excel-power-query-editor.watchFile` — syncs this `.m` every time you save it.

### Toggle Watch

`excel-power-query-editor.toggleWatch` — turns watching on or off for the current file.

### Stop Watching File

`excel-power-query-editor.stopWatching` — stops watching one file, leaving others alone.

## Maintenance

### Cleanup Old Backups

`excel-power-query-editor.cleanupBackups`

Deletes backups beyond the retention limit now, rather than waiting for the next sync to do it.
Respects `backup.maxFiles`.

### Install Excel Symbol Definitions

`excel-power-query-editor.installExcelSymbols`

Registers the Excel symbols with the Power Query language service and reports whether it worked.
This happens automatically at startup, so the command is mainly useful if you installed the Power
Query extension **after** this one and want to check. See [Excel Symbols](Excel_Symbols.md).

### Settings

`excel-power-query-editor.openSettings` — opens VS Code settings filtered to this extension.

## Diagnostics

### Raw Excel Extraction (Debug)

`excel-power-query-editor.rawExtraction`

Dumps what the extension sees inside a workbook, without interpreting it. For attaching to a bug
report when extraction produces something unexpected. Not part of the normal workflow.

---

Settings are in [Config Reference](Config_Reference.md).
