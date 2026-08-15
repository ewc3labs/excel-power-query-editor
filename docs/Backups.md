# Backups

What is kept before a workbook is written, and for how long.

Syncing rewrites a real file that probably represents real work. Before it does, a copy is taken. It
is on by default and there is no reason to turn it off.

## What a backup is

A byte-for-byte copy of the workbook, made immediately before the sync, named after the original:

```
Sales.xlsx
Sales.xlsx.backup.2026-08-15T14-30-00-000Z
```

The extension is not what you recover with. You recover by renaming the file back, in any file
manager. That is deliberate — a backup that needs the tool that made it is not a backup.

## Where they go

```jsonc
"excel-power-query-editor.backup.location": "sameFolder"
```

| Value | Where |
| --- | --- |
| `sameFolder` (default) | beside the workbook |
| `tempFolder` | your system temp directory, under `excel-pq-backups` |
| `custom` | wherever `backup.customPath` points |

`sameFolder` keeps the backup next to the thing it protects, which is where you will look for it.
Choose `tempFolder` if the workbook lives somewhere that syncs — OneDrive and SharePoint will happily
upload every backup you make, and version-control users generally do not want them in the diff.

`custom` takes an absolute path, or one relative to the workspace root:

```jsonc
"excel-power-query-editor.backup.location": "custom",
"excel-power-query-editor.backup.customPath": ".backups"
```

## How many are kept

```jsonc
"excel-power-query-editor.backup.maxFiles": 5
```

Five per workbook. When a sixth is made, the oldest is deleted. The count is **per workbook**, not
per folder, so a busy file cannot push another file's history out.

Cleanup runs after each sync. To run it now instead, use **Cleanup Old Backups** from the palette
(see [Commands](Commands.md)). To keep everything and manage it yourself:

```jsonc
"excel-power-query-editor.backup.autoCleanup": false
```

Be aware of what that means: workbooks are not small, and nothing will ever remove one again.

## When no backup is taken

- **Live sync.** [Live Sync](Live_Sync.md) never writes the file on disk — it asks Excel to change
  the queries in memory. There is nothing to back up, because nothing on disk is touched. Your
  protection there is Excel's own unsaved-changes state: close without saving and the edit is gone.
- **Extraction.** Reading a workbook cannot damage it.
- **`backup.autoBackupBeforeSync` set to `false`.** Then you have said you do not want them.

## Turning it off

```jsonc
"excel-power-query-editor.backup.autoBackupBeforeSync": false
```

Reasonable if the workbook is in version control and you would rather recover from there. Not
reasonable otherwise. A failed sync on a workbook holding six months of query work is not a
hypothetical, and the cost of the setting being on is a few megabytes.

---

Settings are in [Config Reference](Config_Reference.md).
