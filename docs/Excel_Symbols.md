# Excel Symbols

IntelliSense for `Excel.CurrentWorkbook()`.

## The gap this fills

Power Query language support in VS Code comes from Microsoft's
[powerquery.vscode-powerquery][powerquery-vscode] extension. Its standard library is the one Power
Query has in every host, and it contains `Excel.Workbook` — read *an* external workbook file.

It contains no `Excel.CurrentWorkbook`. That function reads *this* workbook, and it is the first
line of a large share of all the M ever written in Excel. It is missing because it does not exist
outside Excel, and the shared library cannot assume Excel.

So the one function you use constantly is the one with no completion, no signature, and no hover.
This extension supplies it.

## What you get

Install both extensions and it happens at startup. Nothing to configure.

Type `Excel.` in a `.m` file and `Excel.CurrentWorkbook` is offered, with its signature and a
description of what it returns.

## Checking it worked

Run **Install Excel Symbol Definitions** from the palette. It registers the symbols and reports what
happened — including which of the two extensions is missing, if one is.

The usual reason for it not working is installation order: if you installed the Power Query
extension *after* this one, this one had nothing to register with at the time. That is handled
automatically — the extension watches for the Power Query extension appearing and registers then —
but the command tells you where things stand without restarting anything.

If the Power Query extension is not installed, nothing is registered and nothing breaks. Everything
else in this extension works without it; you just edit M without language support, which is how it
worked before you had either extension.

## How it works

The Power Query extension exposes `addLibrarySymbols` and `removeLibrarySymbols` on its exported
API. The symbols are handed over in memory, and removed cleanly when this extension deactivates.

**Nothing is written to your workspace and no other extension's settings are modified.** That is
worth stating because it used to be untrue: earlier versions wrote `excel-pq-symbols.json` into
`.vscode/` and appended the folder to `powerquery.client.additionalSymbolsDirectories`. If you have
either of those left over, both are unused and can be deleted — see [Config
Changes](Config_Changes.md).

## Adding more

The definitions are in `resources/symbols/excel-pq-symbols.json`, in the format the Power Query
extension's API expects. If there is an Excel-only function you want completion for, that file is
where it goes and a pull request is welcome.

The right long-term home for these is Microsoft's library rather than ours — see
[vscode-powerquery#206][vscode-powerquery]. Until then, we push in the missing ones.

## Leftovers from the old version

Before 0.7.0 this extension wrote `excel-pq-symbols/excel-pq-symbols.json` to disk and appended that
folder to `powerquery.client.additionalSymbolsDirectories`:

```text
%APPDATA%/Code/User/excel-pq-symbols/          (user scope)
<workspace>/.vscode/excel-pq-symbols/          (workspace scope)
```

**Upgrading does not delete either, deliberately.** An upgrade removing files you might have edited
or moved is a trade nobody agreed to. The extension reports what it finds, once, and leaves the
decision to you.

**If that folder is still listed in the setting, it is not merely untidy** — the Power Query
extension keeps loading that copy from disk while the current symbols arrive through the API. The
file on disk never updates again, so it can only drift. Removing the setting entry is the part worth
doing; the folder itself is harmless once nothing points at it.

## If the symbols stop appearing

They are handed to the Power Query extension through `addLibrarySymbols`, a method described in
[vscode-powerquery#206][vscode-powerquery] rather than in a published API. **We cannot detect the
difference between that method being renamed upstream and the Power Query extension simply being
old** — both look like "the method is not there."

So if completions vanish for someone on a **current** Power Query extension, and
`Excel Power Query Editor: Register Excel Symbols` reports *"too old to accept extra symbols"*,
check the method name upstream before anything else. That message is a guess about the cause, and it
is the wrong guess in exactly this case.

Everywhere the code depends on that method is greppable as `vscode-powerquery#206`.

---

Commands are in [Commands](Commands.md).

[powerquery-vscode]: https://marketplace.visualstudio.com/items?itemName=PowerQuery.vscode-powerquery
[vscode-powerquery]: https://github.com/microsoft/vscode-powerquery/issues/206
