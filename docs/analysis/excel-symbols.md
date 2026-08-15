# Excel symbols and the Power Query language service

**Question:** does Microsoft's Power Query / M extension now recognize the Excel functions on its
own, so we can stop shipping `excel-pq-symbols.json`?

**Answer, measured 2026-08-14 against `powerquery.vscode-powerquery` 1.0.0:** no.

## What its standard library contains

Searching the bundled language server (`server/dist/server.js`, 2.3 MB):

```
Excel.Workbook          present, with documentation, in the embedded library JSON
Excel.CurrentWorkbook   0 occurrences
CurrentWorkbook         0 occurrences
```

So it knows `Excel.Workbook` — read *an* external workbook file — and does not know
`Excel.CurrentWorkbook`, which reads *this* workbook and is where nearly every in-workbook query
begins. Without the extra symbols, that first line lints as an unknown identifier, on essentially
every real file this extension opens.

Our file supplies exactly two symbols: `Excel.CurrentWorkbook` and `Documentation`. Small, and still
load-bearing.

## How the symbols get loaded

The PQ extension exposes `powerquery.client.additionalSymbolsDirectories` — a list of folders it
reads symbol JSON from. That is a clean, supported hook and nothing here needs to change about it.

## The defect: where we put the file

The current `installExcelSymbols` requires an open workspace and writes to:

```
<workspace>/.vscode/excel-pq-symbols/excel-pq-symbols.json
```

Two problems, and the second is the one that matters:

1. **It fails outright with no workspace.** Opening a plain folder - or a single `.m` file, which is
   exactly how someone tries this extension for the first time - throws.
2. **It writes into somebody else's repository.** `.vscode/` is version-controlled in most projects.
   Auto-installing on activation means the extension creates a file in a repo the user did not ask
   us to touch, which may then be committed, reviewed, and wondered about by their colleagues.

## There is an API for this, and it is the answer

`microsoft/vscode-powerquery#206` points at it: the Power Query extension exposes a public API for
exactly this, which the PQ SDK itself uses to push in connector symbols.

```ts
export interface PowerQueryApi {
    readonly onModuleLibraryUpdated: (workspaceUriPath: string, library: LibraryJson) => void;
    readonly addLibrarySymbols: (librarySymbols: ReadonlyMap<string, LibraryJson>) => Promise<void>;
    readonly removeLibrarySymbols: (librariesToRemove: ReadonlyArray<string>) => Promise<void>;
}
```

So the whole problem dissolves. Hand it a `Map` of our two symbols at activation and remove them on
deactivate. **No file is written, no workspace is required, and no other extension's settings are
edited.** All three objections below disappear at once, and 156 lines of file-and-setting machinery
went with them.

Our JSON needed no conversion: checked against the entry the language server holds for
`Excel.Workbook`, the shape is already right — `completionItemKind` and `type`, not the
`completionItemType`/`dataType` an older revision of the API declaration suggested. The data was
always correct; only the delivery was wrong.

## The old design, and why it was wrong



Make it a **command** rather than something that happens to you:

- **"Excel Power Query: Add Excel Symbols File"**, invoked when the user wants it
- Ask where to save it, defaulting somewhere that is *theirs* - global storage, or a path they pick
- Then point `powerquery.client.additionalSymbolsDirectories` at wherever it landed
- Retire the auto-install on activation, and the workspace-relative default with it

The extension's job is to offer the symbols and wire up the setting. Choosing a location in the
user's own project is the user's decision, not ours.
