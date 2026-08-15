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

**The document declares its own scope, and that declaration is the authority.** Not a setting: scope
is a property of the file, and a user with a hundred toolkit workbooks should not be flipping a
global preference between operations. A setting only enables the feature and the extra menu command.

### The manifest

```
// Queries: ALL
```

```
// Queries:
//   fGetNamedRange
//   Sales, EMEA
//   FinalTable
```

`ALL` is case-insensitive, so `all` and `All` are equally valid for hand-rolling.

**A block list rather than a CSV**, for two reasons that only show up in real data:

- **M query names can contain commas.** `#"Sales, EMEA"` is a legal name, and a comma-separated
  manifest cannot represent it without quoting rules that people writing by hand will get wrong.
  One name per line has no separator to escape.
- **A hundred names on one line is unreadable**, and word wrap is off by default in VS Code. A
  manifest nobody can read is a manifest nobody will maintain.

### What each scope licenses

| manifest | update | add | remove |
| --- | --- | --- | --- |
| `ALL` | yes | yes | **yes** — the document is the whole truth |
| a list | yes | yes | **only queries named in the list** |
| absent | yes | yes | **yes** — see below |

### Absent means ALL, and that is not a compromise

The obvious instinct is that an unrecognised document should refuse to delete anything. It is wrong
here, for a reason that only matters because this extension already has users:

**Every `.m` file in existence right now has no manifest**, and today's file sync already replaces
the whole document. Treating "absent" as anything other than `ALL` changes the behaviour of every
existing file on its next sync — either silently, which is worse, or by refusing, which breaks a
workflow thousands of people use daily.

So absent means `ALL`, which is exactly what those files already do. The manifest is how you ask for
*less* authority than the default, and adding one is opt-in.

### Menu commands, not a mode

- **Extract All Power Queries from Excel** — writes `Queries: ALL`
- **Extract Selected Power Queries from Excel** — a picker; writes the block list

Two commands rather than a mode, so the common case costs nobody an extra decision, and the manifest
that lands in the file is a consequence of the command they chose rather than a setting they have to
remember.

## Edge cases worth deciding before building

### A stale `ALL` deletes queries the user has never seen

Extract with `Queries: ALL`. Somebody adds a query in Excel next week. Sync: the document claims to
be the whole truth, so the new query is deleted — and it was never in VS Code, so its removal is
invisible in the diff the user reviewed.

This is the strongest argument for confirming removals (`PQ-25`), and it says something about the
wording: the prompt must name what is about to be **removed**, not summarise what will be written.
"Sync 29 queries?" hides it. "Remove 1 query: NewThing?" does not.

### The manifest and the document can disagree

`Queries: A, B` but the document only contains `shared A`. Two readings: the user deliberately
deleted B, or the manifest is stale.

**Rule: the document says what to write, the manifest says what may be removed.** B is named in the
manifest and absent from the document, so B is deleted. That makes deleting a query the same gesture
as deleting it from the file, which is what someone editing M in an editor expects — and it means
the manifest never needs hand-editing to remove something.

The reverse - `shared C` present but C not named in the manifest - is a query being **added**, which
is always allowed and never needs authority.

### Renaming is a delete plus an add

Rename `shared A` to `shared A2` and the document no longer contains A. Under `ALL`, or under a list
naming A, that removes A and adds A2 — losing anything Excel held against A that is not in the M,
and breaking any query referencing A by name.

This is correct and it is also surprising. Worth a specific line in the confirmation when a removal
and an addition happen in the same sync, because that pattern is far more often a rename than two
separate intentions.

### `ALL` on a workbook with queries that cannot round-trip

If any query in the workbook fails to extract - an unsupported construct, a corrupt DataMashup - and
the document then claims `ALL`, syncing deletes what could not be read. **Extraction must therefore
refuse to write `Queries: ALL` if it did not successfully extract every query it saw**, and downgrade
to a list of the ones it got. A partial extract that lies about being complete is the worst failure
available here.

### Case, whitespace, and hand-rolling

`Queries:ALL`, `queries: all`, `// Queries : ALL` should all work. People will type all of them, and
being strict buys nothing. The one thing worth rejecting is a manifest that parses to zero named
queries while not saying `ALL` - that is a typo, not an instruction, and it should say so rather than
silently authorising nothing.

## Slices

- `PQ-22` — extraction header records scope; the writer reads it
- `PQ-23` — `sync.authority`, honoured identically by both write paths
- `PQ-24` — selective extract: choose queries, and record the subset
- `PQ-25` — confirmation before removing queries, with a setting to suppress it

`PQ-22` gates the rest: without a scope in the document, no other rule here can be enforced safely.
