# Live Sync

Writing Power Query into a workbook Excel already has open.

**Windows and Excel only. Off by default. Beta** — see [What has not been
tried](#what-has-not-been-tried).

## The problem it removes

Excel holds an open workbook with an exclusive write lock. Editing a `.m` file and syncing while the
workbook is open used to produce *"the file is locked, close Excel and try again"* — so you closed
the workbook, synced, reopened it, found your place again, and lost your thread.

Live sync does not fight the lock. It asks Excel to make the change through Excel's own object
model, so the file on disk is never touched.

## Turning it on

```jsonc
"excel-power-query-editor.sync.liveWhenOpen": true
```

Then sync as normal. The extension decides per file, every time:

| | |
| --- | --- |
| Workbook **closed** | the file is rewritten directly, exactly as before |
| Workbook **open in Excel** | the change goes through Excel |
| Excel not running, or not this workbook | the file is rewritten directly |

There is nothing else to configure and no mode to remember being in.

## What happens when it works

The queries in the open workbook are updated in place, and **Excel shows unsaved changes** — the
same state as if you had typed the edit yourself. You save when you are ready, in Excel.

That is the part worth internalizing: **your file on disk has not changed.** If you close the
workbook without saving, the edit is gone from Excel and still present in your `.m` file.

Three behaviors worth knowing:

- **Queries you did not touch are left alone.** A query whose M is already identical is not
  rewritten, so the workbook is not marked as modified for no reason.
- **A query in the workbook that is not in your `.m` file is reported, never deleted.** The log
  names it. Nothing is removed.
- **New queries are added.** A `shared` binding with no matching query creates one.

## What it handles

Measured rather than assumed:

- **Several Excel instances at once.** Workbooks are located through the Running Object Table, so it
  does not matter which instance holds yours. This is the case that defeats most Excel automation,
  and the same split that stops you copying between two open workbooks.
- **OneDrive and SharePoint.** A synced workbook is registered by Excel under its cloud URL rather
  than the local path you see, and matching accounts for that.
- **Excel being busy.** Modal dialogs, recalculation and cell-edit mode all cause COM calls to fail.
  Those are transient and are retried rather than reported as errors.
- **A workbook that is not open.** Checking never opens anything.

## What has not been tried

Untested, not known-broken:

Protected View · workbooks opened from a web link · co-authored files with AutoSave · `excel.exe /x`
· environments where group policy restricts automation · Office update channels other than Current.

The worst plausible failure is that live sync declines and the normal file path takes over — which
is what would have happened anyway. It cannot corrupt a workbook, because it never writes one.

## When it does not work

Set the log level and look at one line:

```jsonc
"excel-power-query-editor.log.level": "debug"
```

**View → Output → Excel Power Query Editor**, then find:

```
Excel file is locked; live sync CAN/cannot handle it (available=…, excelProcesses=…)
```

That single line says which branch was taken and why. Common answers:

| What you see | What it means |
| --- | --- |
| `live sync CAN handle it` | it worked, or the failure is later — read on in the log |
| `available=false, not-windows` | live sync needs Windows |
| `Excel is running but this workbook is not visible` | usually VS Code and Excel at different elevation levels — COM hides running objects across integrity levels. Run both normally |
| `Workbook is not open in Excel` | the workbook really is not open. Check you are syncing the `.m` that matches the workbook you have open |

If it is none of those, [open an issue][open-an-issue] with that line.

## How it works

Excel exposes Power Query through its object model: `Workbook.Queries` is a collection, and each
query's `Formula` property — the M — is readable and writable. A helper script attaches to the
running Excel, finds the workbook, and sets the formulas.

The design, including the approaches that were rejected and why, is in
[design/live-sync-to-open-excel.md][design-live-sync-to].

[design-live-sync-to]: design/live-sync-to-open-excel.md
[open-an-issue]: https://github.com/ewc3labs/excel-power-query-editor/issues
