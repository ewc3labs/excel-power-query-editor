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

## Save the workbook first

If the workbook has unsaved changes in Excel, live sync **declines** and asks you to save it.

That looks like an inconvenience and is the one place where a backup cannot help you. The backup
taken before a sync is a copy of the file **on disk** - the last saved state. Edits sitting unsaved
in Excel are in no file anywhere. Writing over them would destroy the only copy, while the backup
sat there looking like protection.

Saving on your behalf would be worse: it commits changes you may still have been deciding about. So
the extension says what is in the way and leaves the decision where it belongs.

Save in Excel, then sync. Once the workbook is clean, the backup genuinely covers the state you are
about to change.

### Turning it off

```jsonc
"excel-power-query-editor.sync.requireSavedWorkbook": false
```

There is a real workflow this gets in the way of: hammering on prototype queries in a scratch
workbook, where saving between every sync is friction for no benefit. If that is what you are doing,
turn it off.

Be clear about the trade. With it off, a sync you did not want cannot be undone from a backup - the
backup holds the last state **you** saved, not the state you were working in. The extension writes a
warning to the log each time it proceeds over unsaved changes, so a surprising result later has an
explanation.

Closed workbooks are unaffected either way: they are written to disk with a backup, as always.

### AutoSave has not been tested

Workbooks in OneDrive and SharePoint usually have **AutoSave on**, and Excel then persists changes
continuously. What that does to the "review before saving" promise above is currently **unknown** -
a live write may be saved to the cloud within seconds, with no review step.

It may be perfectly fine, and version history may cover you. It has not been measured, so it is
listed here rather than quietly assumed. If you rely on being able to close without saving, check
that AutoSave is off for that workbook.

## Excel and VS Code must run at the same level

**This is a Windows limitation, not a missing feature.** COM partitions its running object table by
integrity level: a workbook open in an elevated Excel is invisible to a normal VS Code, and a
workbook open in a normal Excel is invisible to an elevated VS Code. Neither side can see across,
and no amount of code on our part changes that.

What the extension does about it is refuse to guess. It can prove the file is locked - the write
fails - but if it cannot reach whatever holds the lock, it says so and stops. It will **not** fall
back to a similarly named workbook it can see. That fallback used to exist, and on a real machine it
matched a workbook two directories away with the same name and would have written the wrong queries
into it.

Two ways forward, both yours to pick:

- **Run VS Code and Excel the same way**, normally for preference. Then live sync works.
- **Close the workbook.** With it closed the file is written directly, and none of this applies.

## Cloud workbooks are matched exactly

A workbook synced by OneDrive or SharePoint is registered by Excel under its **cloud URL**, not the
local path you see. The extension translates your local path to that URL using OneDrive's own
sync-root registration, then matches the two **exactly**.

There is deliberately no fuzzy matching. Comparing trailing path segments seems reasonable and is
not safe: `...\Documents\PowerQuery\Book.xlsx` and `.../Scripts/PowerQuery/Book.xlsx` agree on two
segments while being entirely different files. A false negative here means live sync declines and
you close the workbook. A false positive means your queries land in somebody else's workbook.

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
