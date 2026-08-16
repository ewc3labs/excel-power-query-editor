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
Choose `tempFolder` if the workbook lives somewhere that syncs — OneDrive and SharePoint will
happily upload every backup you make, and version-control users generally do not want them in the
diff.

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

- **Extraction.** Reading a workbook cannot damage it.
- **`backup.autoBackupBeforeSync` set to `false`.** Then you have said you do not want them.

## Live sync still takes one

[Live Sync](Live_Sync.md) does not write the file on disk, so it is tempting to think a backup is
pointless. **A backup is taken anyway, deliberately.**

A live write leaves the workbook open and modified in Excel. The moment you press save - which is
the whole point of the feature - Excel writes over the file on disk. If that is the first time the
file has changed since you started editing, the version you had is gone and nothing preserved it.

So the backup protects the on-disk state you may be about to overwrite from inside Excel, not the
write the extension performs. You will see backup files and retention cleanup during live sync, and
that is correct.

## Turning it off

```jsonc
"excel-power-query-editor.backup.autoBackupBeforeSync": false
```

Reasonable when something else is already keeping your history:

- **The workbook is in OneDrive or SharePoint.** Both keep version history, which recovers the whole
  workbook rather than just the copy we happened to take, and does it without filling a synced
  folder with `.backup.` files. If you are working entirely in cloud storage, this is a fair trade.
- **The workbook is in version control.**

Not reasonable otherwise. A failed sync on a workbook holding six months of query work is not a
hypothetical, and the cost of leaving it on is a few megabytes.

If the churn is what bothers you rather than the safety, turn down `backup.maxFiles` or point
`backup.location` at `tempFolder` instead of switching backups off. Retention exists precisely
because a fast edit-sync loop was producing more backups than anyone wanted.

---

Settings are in [Config Reference](Config_Reference.md).
