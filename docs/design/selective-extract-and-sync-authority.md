# Selective extract, and who is authoritative

**Status:** design. Slices `PQ-22` … `PQ-25`. Nothing here is built.

## The asymmetry we already shipped

The two write paths have **opposite semantics**, and neither is labelled:

| | a query in the workbook but NOT in the `.m` file |
| --- | --- |
| **File sync** (`setFormula` with the whole section document) | **deleted** |
| **Live sync** (per-query diff through Excel) | **reported, never touched** |

Same command, same file, opposite outcome — decided by whether Excel happens to have the workbook
open at that moment. A user who learns one behaviour will be surprised by the other, and the
surprise is silent in both directions: one deletes work they meant to keep, the other leaves a query
they meant to remove.

This was not a decision. The file path predates live sync and replaces the whole DataMashup because
that is what `excel-datamashup` offers; live sync works per query because that is what the Excel
object model offers. Each is the natural shape of its own mechanism, and nobody chose the difference.

## Selective extract turns a surprise into a hazard

The proposal: extract only chosen queries. One folder of 100 toolkit queries, pull one into VS Code
to fold into another workbook.

That is clearly useful, and it is **incompatible with whole-document replace**. Extract 1 of 29
queries, edit it, press sync with the file closed, and today's file path replaces the entire section
document with your one query. The other 28 are gone — from a workbook the user never intended to
change beyond one function.

**So the rule the feature needs is:**

> A partial extract must never license a full replace.

Which means the `.m` file has to know what it is. A document holding every query and a document
holding one query look identical to the writer, and only one of them may be treated as the truth
about the whole workbook.

## The design

### 1. The `.m` records its own scope

The extraction header already exists and is already ignored by the writer, which strips everything
above `section`. Make it machine-readable:

```
// Power Query extracted from: Toolkit.xlsx
// Extracted: all (29 queries)
```

```
// Power Query extracted from: Toolkit.xlsx
// Extracted: subset (2 of 29) — fGetNamedRange, FinalTable
```

A missing header means "unknown", and unknown is treated as a subset — the conservative reading,
because a hand-written `.m` is far more likely to be a fragment than a faithful copy of a workbook.

### 2. One setting, one meaning, both paths

```jsonc
"excel-power-query-editor.sync.authority": "merge" | "replace"
```

- **`merge`** *(proposed default)* — update what matches, add what is new, **never remove**. Today's
  live behaviour.
- **`replace`** — the document is the truth: queries absent from it are removed. Today's file
  behaviour, and what mass-modify-and-tear-out workflows want.

Applied to **both** paths, so the answer no longer depends on whether Excel is open. Live sync can
honour `replace` — `WorkbookQuery` exposes `Delete()` — and the file path can honour `merge` by
splitting the existing section, applying the incoming queries over it, and writing the union back.

### 3. `replace` on a subset is refused

The two rules together:

| document scope | `merge` | `replace` |
| --- | --- | --- |
| **all** | update + add | update + add + **remove missing** |
| **subset** | update + add | **scoped to the named queries only** — never removes anything outside them |
| **unknown** | update + add | **refuse, and say why** |

A subset under `replace` does the honest thing: it is authoritative *over the queries it names* and
silent about the rest. That is what somebody pulling one query out of a hundred means, and it is
never what "delete the other 99" means.

### 4. Removal is confirmed, once

Whatever the mode, deleting queries is the one action here that destroys work not recoverable from
the `.m` file. It gets a confirmation naming the queries and the count, in the same spirit as the
backup that already happens before every write. A setting can suppress it for people who do this
all day, and the default is to ask.

## Why the default should be `merge`

It is the safer half of a behaviour users cannot currently predict, and the cost of being wrong is
asymmetric: `merge` leaves a stale query in a workbook, which is visible and fixable; `replace`
deletes work, which may not be noticed until the workbook is opened weeks later.

Changing the file path's default from today's implicit `replace` is a behaviour change and belongs
in a minor version with a changelog entry that says so plainly.

## Slices

- `PQ-22` — extraction header records scope; the writer reads it
- `PQ-23` — `sync.authority`, honoured identically by both write paths
- `PQ-24` — selective extract: choose queries, and record the subset
- `PQ-25` — confirmation before removing queries, with a setting to suppress it

`PQ-22` gates the rest: without a scope in the document, no other rule here can be enforced safely.
