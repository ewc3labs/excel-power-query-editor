# Watch Mode

Syncing automatically every time you save.

The normal loop is edit, save, run **Sync Power Query to Excel**. Watch mode removes the third step:
save the `.m` file and the workbook is updated. Nothing else changes — the same sync runs, with the
same backup and the same choice between writing the file and going through Excel.

## Turning it on for one file

Right-click a `.m` file and choose **Watch File for Changes**, or use **Toggle Watch** from the
palette. Watching is per file, and survives until you stop it or close the window.

**Stop Watching File** stops that one file. Others carry on.

## Turning it on for everything

```jsonc
"excel-power-query-editor.watch.always": true
```

Every file you extract is watched from the moment it is created. This is the setting to use if watch
mode is simply how you work, and it is the one most people end up on.

It has a deliberate ceiling:

```jsonc
"excel-power-query-editor.watch.maxFiles": 25
```

Twenty-five watchers. A workspace with two hundred `.m` files in it would otherwise put two hundred
file watchers into VS Code, and you would feel it. When the limit is reached, no more are added and
the log says so.

## What happens on save

1. The file is written to disk by VS Code, as normal.
2. The sync runs.
3. If the workbook is open in Excel and [Live Sync](Live_Sync.md) is on, the change goes through
   Excel. Otherwise the file is rewritten and a [backup](Backups.md) is taken first.

Watch mode does not batch or debounce beyond what the editor does. One save is one sync.

## When the workbook is locked

```jsonc
"excel-power-query-editor.watch.checkExcelWriteable": true
```

On by default. Before writing, the extension checks whether the workbook can actually be written and
reports it clearly instead of failing part-way through. With live sync enabled this rarely comes up,
because an open workbook is handled through Excel rather than through the lock.

## Deleting a watched file

```jsonc
"excel-power-query-editor.watch.offOnDelete": true
```

On by default: delete the `.m` and it stops being watched. Turn it off only if you regularly delete
and recreate the same path and want the watcher to survive it — for example if a build step
regenerates your `.m` files.

## Turning it off

**Toggle Watch** on the file, or set `watch.always` to `false` for new extractions. Closing the
window ends all watching; nothing persists across sessions.

---

Commands are in [Commands](Commands.md), settings in [Config Reference](Config_Reference.md).
