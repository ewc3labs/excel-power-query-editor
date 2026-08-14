# Live sync to an open workbook

**Status:** mechanism proven 2026-08-14. Not built. Slices `PQ-12` … `PQ-15`.

## The ask

[Discussion #3](https://github.com/ewc3labs/excel-power-query-editor/discussions/3) /
[issue #4](https://github.com/ewc3labs/excel-power-query-editor/issues/4), from **namgaw (James)**,
October 2025:

> I was hoping that this extension would enable me to writeback to Excel while the file is open —
> basically replacing the default "advanced" editor.

He had been using **Monkey Tools**, Ken Puls' Excel add-in, which can do this, and lost access when his
employer locked the network down. He suggested contacting Ken Puls.

This is the single most valuable thing the extension does not do. Everything about the current design —
"close the file so we can write it", the watch-and-wait timer, `watch.checkExcelWriteable` — exists to
work around the fact that Excel holds an open workbook exclusively. Live sync does not work around that
problem. **It stops having it.**

## Why the lock is not the fight

The instinct is to make the *file* writable while Excel holds it. That is the wrong target and it is
not winnable — and even if it were, Excel would overwrite whatever we wrote when the user saved.

Excel exposes Power Query through its object model. `Workbook.Queries` (Excel 2016+) is a collection of
`WorkbookQuery`, and **`WorkbookQuery.Formula` — the M code — is read/write**. So the supported route is
not the file at all. It is the running application.

No RPA. No driving the Advanced Editor UI. No zip surgery on a locked file.

## Proven, on this machine

Measured rather than assumed, against `test/fixtures/simple.xlsx`:

```powershell
$wb = $xl.Workbooks.Open($path)
$wb.Queries.Count            # 1
$wb.Queries.Item(1).Name     # StudentResults
$wb.Queries.Item(1).Formula  # let Source = Excel.CurrentWorkbook(){[Name="StudentNames"]}[Content] ...
$wb.Queries.Item(1).Formula = $newM     # accepted
```

And then the case that actually matters — a workbook **the user already has open**, reached from a
separate process the way an external tool must:

```
attached to running Excel: 16.0
found open workbook: livesync-test.xlsx  (saved=True)
after write, first line: // edited live ... 23:39:21
WRITE TO OPEN WORKBOOK: SUCCESS
workbook now dirty (needs save): True
```

`Marshal.GetActiveObject("Excel.Application")` finds the running instance through the Running Object
Table; the workbook is located by name; the formula is rewritten in place. The workbook goes dirty
exactly as if the user had typed it, and the user saves when they choose.

**Excel, unlike OneNote, still serves external automation.** That is not a given — the sibling project
`ewc3-recall-tape` exists because OneNote stopped doing so — and it is why this is worth building.

## Architecture

Two ways in, and the second is not needed first.

### A. VS Code → COM, via a spawned helper *(start here)*

The extension shells out to a small helper that does the COM work and returns JSON. A first version can
be a PowerShell script, which adds **no dependency at all** — an important property for a project whose
README promises "Zero Dependencies" and cross-platform support.

- Cheapest possible first cut; the whole mechanism is four object-model calls.
- Native COM bindings (`winax`) were rejected for v1: a native module means prebuilds per Node ABI,
  breaks the zero-dependency claim, and complicates CI for a feature only Windows users can run.
- A compiled C# helper is the natural upgrade if PowerShell start-up cost or execution policy bites.

### B. An Excel-side add-in talking to VS Code

Wilson's original instinct, and the ShowCase/Essbase shape: a COM/VSTO add-in inside Excel with a local
channel to the extension. It is how Monkey Tools reaches the object model, and it is exactly the
architecture already proven in `ewc3-recall-tape` (protocol handler → named pipe → in-process add-in).

Worth it only if automation turns out to be blocked in locked-down environments — which is precisely
what happened to James. Note that **Office JS add-ins cannot do this**: that API has no Power Query
surface at all. It would have to be VSTO/COM.

## Constraints, stated plainly

- **Windows and Excel only.** The extension's core promise is editing `.xlsx` with no Excel and no
  Windows. Live sync is an *additional mode* for people who have both, never a replacement, and the
  docs must not blur that.
- **Elevation must match.** A non-elevated helper cannot see an Excel started elevated, or vice versa.
- **Multiple Excel instances** — `GetActiveObject` returns one. Workbook lookup must handle "open in a
  different instance" rather than assuming.
- **The workbook goes dirty.** We should almost certainly *not* save on the user's behalf; making
  something dirty is honest, saving for them is not.
- **Refresh is a separate decision.** Setting `Formula` changes the query; it does not re-run it.
  Whether to refresh, and whether that is a setting, is a UX question not a technical one.

## The two writers do not write the same thing

This is the finding that shapes the whole feature, and it is not visible until you look at both
sides.

**On disk, a `.m` file is a SECTION DOCUMENT holding every query in the workbook:**

```
section Section1;

shared fGetNamedRange = let ... ;
shared RawInput       = let ... ;
shared FinalTable     = let ... ;
```

`test/fixtures/expected/complex_FinalTable.m` has three `shared` blocks in one file. Extraction pulls
**all** queries out together, and the disk writer puts them **all** back together.

**Through COM, `WorkbookQuery.Formula` is ONE query's expression, with no wrapper:**

```
let
    Source = Excel.CurrentWorkbook(){[Name="StudentNames"]}[Content],
    ...
in
    Source
```

Measured: 160 characters written through COM came back from disk as 199. The 46-character delta is
exactly `section Section1;`, the blank line, `shared StudentResults = `, and the closing `;`.

So the live path cannot simply hand the file's contents to Excel. It has to:

1. Parse the section document into `{ name -> expression }`
2. Walk `Workbook.Queries` and set `Formula` per matching name
3. Decide what to do about queries present on one side and not the other — a `shared` block with no
   matching query means `Queries.Add`; a query with no matching block means the user deleted it, and
   deleting someone's query is not something to do on a guess

**Zero drift is achievable, but the invariant has to be stated as a round trip, not a byte compare:**

> section document → split → N × `Formula` → save → extract → section document

must return the original text. Anything else — reordering, whitespace normalisation, a lost trailing
semicolon — is drift, and drift means two sources of truth in a tool whose whole job is not corrupting
workbooks. That is `PQ-15`, and it should be a test with real fixtures before `PQ-13` ships.

**What does survive, verified:** a formula written through COM and read back from the saved file kept a
literal tab, trailing spaces, all seven CRLFs, and `café naïve` intact. The M itself is not mangled.
The only difference is the wrapper.

## Bulk extraction is untouched, and never needed a close

Extraction reads the file; **Excel takes a write lock, not a read lock.** Measured against a workbook
open in Excel:

```
plain read: OK, 38596 bytes
zip open  : OK, 22 entries, customXml present=True
open for write FAILED: PermissionError
```

So ripping M out of a folder of workbooks works whether they are open or closed, needs no Excel, no
COM and no Windows, and **live sync changes none of it**. Only the write path was ever blocked, and
live sync is a write-path option that appears when Excel happens to be running.

The two-tailed rule is therefore simple to state: **write into the zip when the file is closed, write
through COM when it is open.** One decision, made per file, at write time.

Note the `customXml/item1.xml` part is **UTF-16LE** with a BOM. Read it as a buffer and sniff, which
is what the extension already does; decoding it as UTF-8 yields mojibake and `DataMashupNotFound`.

## Open questions for the slices

1. How does a `.m` file map to a query name — filename convention, a header comment, or a stored map?
2. What happens when the workbook is open but the query does not exist yet — `Queries.Add`?
3. Fall back silently to the existing close-the-file flow when Excel is not running, or say so?
4. Is `Formula` round-trip lossless against what `excel-datamashup` writes on disk? A query written
   live and then read back from the file should be byte-identical, or we have two sources of truth.

## Ken Puls

James's suggestion is a good one and costs nothing. Ken Puls (Excelguru, Monkey Tools, *M is for (Data)
Monkey*) has solved this problem in the other direction — inside Excel, reaching out — and this project
is coming at it from VS Code, reaching in. Credit generously either way; the acknowledgements section
already names the shoulders this stands on.
