<!-- HEADER_TABLE -->
<table align="center">
<tr>
  <td width="112" align="center" valign="middle">
    <img src="assets/excel-power-query-editor-logo-128x128.png" width="128" height="128"><br>
    <strong>E · P · Q · E</strong>
  </td>

  <td align="center" valign="middle">
    <h1 align="center">Excel Power Query Editor</h1>
    <p align="left">
      <b>Edit Power Query M in VS Code — including in workbooks Excel already has open.</b><br>
      <sub>
        No Excel automation to install, no COM to configure, no closing your file to save your work.
        Built by <strong>EWC3 Labs</strong>.
      </sub>
    </p>
  </td>

  <td width="112" align="center" valign="middle">
    <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128"><br>
    <strong><em>QA Officer</em></strong>
  </td>
</tr>
</table>
<!-- /HEADER_TABLE -->

<!-- BADGES -->
<p align="center">
  <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <img alt="Version" src="https://img.shields.io/badge/Version-0.6.0-brightgreen.svg">
  <img alt="Tests" src="https://img.shields.io/badge/tests-115%20passing-brightgreen.svg">
  <a href="https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor"><img alt="VS Code Marketplace" src="https://img.shields.io/badge/VS_Code-Marketplace-blue.svg"></a>
  <a href="https://www.buymeacoffee.com/ewc3labs"><img alt="Buy Me a Coffee" src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-yellow?logo=buy-me-a-coffee&logoColor=white"></a>
</p>
<!-- /BADGES -->

---

Excel's Advanced Editor gives you one query at a time, in a modal dialog, with no version control and
no real editor. This gives you your M code as files — in VS Code, with IntelliSense, multi-cursor,
diff, blame, and everything else you already use.

It reads `.xlsx`, `.xlsm` and `.xlsb` directly. **Excel does not need to be installed**, which also
means it works on macOS and Linux, and in CI.

## New in 0.6.0 — write to a workbook that's open

The oldest complaint about tools like this: you edit your query, hit save, and get *"the file is
locked, close Excel and try again."* So you close the workbook, sync, reopen it, find your place
again, and lose your train of thought.

**Turn on `sync.liveWhenOpen` and that stops happening.** If Excel already has the workbook open, the
change goes straight into it through Excel itself:

- The file on disk is **not touched**. Excel simply shows unsaved changes, exactly as if you'd typed
  them, and you save when you're ready.
- Queries you didn't touch are left alone. Identical M isn't rewritten, so nothing is marked dirty
  for no reason.
- A query in the workbook that isn't in your `.m` file is **reported, never deleted**.
- Workbook closed? It writes to the file as it always has. One decision, made per file, per save.

Live sync needs Windows and Excel, because it works by asking Excel to make the change. Yes: the
extension that famously doesn't require Excel requires Excel to write to Excel while Excel has the
file open. We're at peace with it.

Everything else still needs neither. Extract, edit, bulk-extract, and sync to closed workbooks all
work on macOS and Linux with no Excel installed, exactly as they always have.

> Requested by [@namgaw](https://github.com/ewc3labs/excel-power-query-editor/discussions/3), who
> wanted to stop closing his workbook to save his own query. Fair.

> **Beta.** It works by talking to Excel through COM, and Excel varies enormously in the wild. It
> never writes your file, so the worst case is that it declines and the normal path takes over. See
> the [changelog](CHANGELOG.md) for what was measured to work and what has not been tried — and
> please [report anything odd](https://github.com/ewc3labs/excel-power-query-editor/issues).

## Quick start

**1. Install** — from the
[Marketplace](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor),
or `ext install ewc3labs.excel-power-query-editor`.

**2. Extract** — right-click any `.xlsx` / `.xlsm` / `.xlsb` → **Extract Power Query**. You get a
`.m` file beside it with every query in the workbook.

**3. Edit and sync** — edit the `.m`, then right-click → **Sync to Excel**. Or turn on **Watch** and
it syncs as you save.

That's the whole workflow. Nothing to configure to get started.

## What it does

| | |
| --- | --- |
| **Edit M as files** | Full VS Code editing, IntelliSense, and your own git history |
| **No Excel required** | Reads the file format directly — Windows, macOS, Linux, CI |
| **Excel required** | ...only for live sync, obviously. It's in the name |
| **Live sync** | Write into a workbook Excel has open, without closing it *(Windows)* |
| **Watch mode** | Sync on save, with a configurable debounce |
| **Automatic backups** | Every write is backed up first, with configurable retention |
| **Bulk extract** | Select a hundred workbooks and pull the M out of all of them |
| **Excel symbols** | `Excel.CurrentWorkbook()` IntelliSense, which the M language service doesn't ship |

## Your data

This extension writes to spreadsheets, and it treats that as the serious thing it is.

- **A backup before every write**, in a location you choose.
- **Never edits in place** — a failed write leaves the original untouched.
- **Live sync doesn't write the file at all.** It hands the change to Excel and leaves the saving to
  you.
- **No telemetry.** Nothing is collected, sent, or phoned home.

If you find a case where a workbook is damaged, that is the highest-priority bug this project can
have. [Open an issue](https://github.com/ewc3labs/excel-power-query-editor/issues) and it goes to the
front of the queue.

## Configuration

Everything works out of the box. Common ones:

```jsonc
{
  // Write into the workbook when Excel already has it open (Windows + Excel)
  "excel-power-query-editor.sync.liveWhenOpen": true,

  // Sync automatically when you save the .m file
  "excel-power-query-editor.watch.always": true,

  // Keep backups next to the workbook, retaining the last 10
  "excel-power-query-editor.backup.location": "sameFolder",
  "excel-power-query-editor.backup.maxFiles": 10
}
```

Full reference: [Config Reference](docs/Config_Reference.md).

> **Upgrading from 0.5.x?** The settings were reorganized into namespaces (`watchAlways` →
> `watch.always`, and so on). **Your existing settings are migrated automatically** the first time
> 0.6.0 starts, and the old names remain documented as deprecated for a release.

## Documentation

- **[Overview](docs/Overview.md)** — the documentation index
- **[User Guide](docs/User_Guide.md)** — the full workflow
- **[Config Reference](docs/Config_Reference.md)** — every setting
- **[Commands](docs/Commands.md)** — every command
- **[Live Sync](docs/Live_Sync.md)** — writing to an open workbook
- **[Config Changes](docs/Config_Changes.md)** — read this when upgrading
- **[Live sync design](docs/design/live-sync-to-open-excel.md)** — how writing to an open workbook
  works, and what was measured to make it reliable
- **[Contributing](docs/CONTRIBUTING.md)** · **[Changelog](CHANGELOG.md)** ·
  **[Support](SUPPORT.md)**

## Acknowledgments

This stands on other people's work:

- **[Vladinator](https://github.com/Vladinator)** —
  [excel-datamashup](https://github.com/Vladinator/excel-datamashup), which does the genuinely hard
  part: reading and writing the DataMashup part inside a workbook.
- **[Alexander Malanov](https://github.com/amalanov)** —
  [EditExcelPQM](https://github.com/amalanov/EditExcelPQM), which showed this was possible.
- **[Microsoft](https://marketplace.visualstudio.com/items?itemName=PowerQuery.vscode-powerquery)** —
  the Power Query / M language extension that provides the language service.
- **[Ken Puls](https://excelguru.ca/)** — whose Monkey Tools proved writing to an open workbook was
  a solved problem, from the other side of the same wall.

---

<p align="center">
  <sub>MIT licensed · Built by <a href="https://github.com/ewc3labs">EWC3 Labs</a> ·
  <a href="https://www.buymeacoffee.com/ewc3labs">Buy me a coffee</a> if it saved you an afternoon</sub>
</p>
