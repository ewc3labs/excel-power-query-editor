# Installation

Installing the extension, and what it needs.

## Requirements

| | |
| --- | --- |
| **VS Code** | 1.101.0 or newer |
| **Excel** | not required — see below |
| **Operating system** | Windows, macOS or Linux |

**Excel does not need to be installed.** Extracting and syncing read and write the workbook file
directly, so the whole core workflow runs on a machine that has never had Office on it. Excel is
required for exactly one feature, [Live Sync](Live_Sync.md), which writes into a workbook Excel
already has open — and that is Windows-only, because it drives Excel through COM.

## From the Marketplace

```
ext install ewc3labs.excel-power-query-editor
```

Or search for "Excel Power Query Editor" in the Extensions view, or use the
[Marketplace listing](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor).

## From a release VSIX

The Marketplace listing currently lags the repository — automatic publishing is deliberately
switched off while the current run of changes settles. **The newest build is on the
[releases page](https://github.com/ewc3labs/excel-power-query-editor/releases)**, as a `.vsix` file.

Download it, then either:

- **In VS Code** — Extensions view → `...` menu → **Install from VSIX...**
- **From a terminal** —
  ```
  code --install-extension excel-power-query-editor-0.6.0.vsix
  ```

If you run VS Code Insiders, install into Insiders explicitly — `code-insiders --install-extension`.
The two keep separate extension sets and separate settings, and installing into one while working in
the other produces a confusing half-hour.

## Recommended companion

Install Microsoft's
[Power Query / M Language](https://marketplace.visualstudio.com/items?itemName=PowerQuery.vscode-powerquery)
extension as well. It provides syntax highlighting, formatting and IntelliSense for the `.m` files
this extension produces, and this extension adds the Excel-specific symbols that it is missing — see
[Excel Symbols](Excel_Symbols.md).

Not required. Without it you edit M as plain text.

## Verifying

Right-click any `.xlsx`, `.xlsm` or `.xlsb` file in the Explorer. **Extract Power Query** should be
on the menu. That is the whole check — if it is there, the extension is loaded.

For anything deeper, open **View → Output → Excel Power Query Editor**, which is where the extension
reports what it is doing.

## Upgrading

Read [Config Changes](Config_Changes.md) before upgrading. Settings do get renamed, and while
migration is automatic, that document is where you find out what happened to a setting that is no
longer where you left it.

## Uninstalling

Uninstall from the Extensions view. Nothing is left behind: no files are written into your
workspace, and no other extension's settings are modified. Your `.m` files and any
[backups](Backups.md) are ordinary files and remain where they are — delete them yourself if you
want them gone.

---

Next: [User Guide](User_Guide.md).
