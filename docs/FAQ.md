# FAQ

The questions that come up most.

### Do I need Excel installed?

No. Extracting and syncing read and write the workbook file directly, so the core workflow runs on
macOS, on Linux, and on a Windows machine that has never had Office on it.

Excel is required for one feature — [Live Sync](Live_Sync.md), writing into a workbook Excel already
has open — and that is Windows-only. No Excel required for file manipulation. Excel required for
live sync to Excel. Naturally.

### Can I extract while the workbook is open?

Yes. Reading a workbook is never blocked, and extraction does not disturb Excel or your file. It is
only *writing* a file Excel holds open that the lock prevents — and [Live Sync](Live_Sync.md)
removes that too.

### Does syncing overwrite queries I did not touch?

No. Queries whose M is unchanged are left exactly as they are.

Under live sync, a query that exists in the workbook but not in your `.m` file is **reported and
never deleted**. The extension will not remove something you did not ask it to remove.

### Where did my setting go?

Every setting moved into a namespace in 0.6.0 — `logLevel` became `log.level`, and so on. **Your
values were migrated automatically**, in whatever scope you had set them.

[Config Changes](Config_Changes.md) has the full old-to-new table.

### Which file does a `.m` sync back to?

The workbook it was extracted from, matched by name in the same folder: `Sales.xlsx` produces
`Sales.m`, and syncing `Sales.m` writes `Sales.xlsx`. Move one without the other and the pairing
breaks.

### What is a `.m` file, exactly?

A **section document** — one file holding every query in the workbook, each as a `shared` binding.
It is not one query. That is why extraction gives you the whole workbook rather than a file per
query, and why syncing considers all of them.

### Can I edit just one query?

Today, a sync considers the whole section document — which is also what makes it predictable. Work
on the one query you care about and leave the rest of the file alone; unchanged queries are not
rewritten, so the effect is the same.

Editing a subset explicitly is designed but not yet built —
[design/selective-extract-and-sync-authority.md][design-selective].

### Is my workbook safe?

The core protection is that a [backup](Backups.md) is taken before every file write, kept five deep
per workbook, and recoverable by renaming it — no tool required.

Live sync is safer still, structurally: it never writes the file at all. It asks Excel to change the
queries in memory, so the workbook shows unsaved changes and you decide whether to keep them.

### Live sync is not doing anything.

Set `log.level` to `debug`, open **View → Output → Excel Power Query Editor**, and find the line
beginning `Excel file is locked; live sync CAN/cannot handle it`. That one line says which branch
was taken and why.

The most common cause is elevation: if VS Code and Excel are running at different privilege levels,
COM cannot see across them. Run both normally. [Live Sync](Live_Sync.md) has the full table.

### Why is `Excel.CurrentWorkbook()` not in IntelliSense?

Because Microsoft's Power Query extension ships the standard library, which has no Excel-only
functions in it. This extension supplies them — see [Excel Symbols](Excel_Symbols.md). If completion
is missing, run **Install Excel Symbol Definitions**, which reports what is actually registered.

### Does it work with OneDrive or SharePoint?

Yes, including live sync — a synced workbook is registered by Excel under its cloud URL rather than
the local path you see, and matching accounts for that.

One suggestion: set `backup.location` to `tempFolder` for cloud-synced workbooks, or every backup
you make gets uploaded.

### Something is wrong and this page did not cover it.

Set `log.level` to `debug`, reproduce it, and [open an issue][open-an-issue] with the output. The
log is written to be readable by whoever is holding the problem, not just by us.

[design-selective]: design/selective-extract-and-sync-authority.md
[open-an-issue]: https://github.com/ewc3labs/excel-power-query-editor/issues
