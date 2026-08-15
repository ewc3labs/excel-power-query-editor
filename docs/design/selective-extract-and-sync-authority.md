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

## `ALL` means exactly what it means today

The single hardest requirement here, and it constrains everything else:

> **Under `ALL`, the file path does what it does now: `setFormula(wholeDocument)`.**

No diffing, no per-query reasoning, no interpretation. Thousands of people rely on "replace the
DataMashup with what is in this file", it is simple and predictable, and every clever thing added to
it is a new way for it to be wrong. The manifest feature is **additive**: it introduces a way to ask
for *less*, and changes nothing about asking for everything.

That kills a piece of the earlier design outright. **There is no rename detection.** Under `ALL`, a
rename is not a rename — it is a document that no longer contains `A` and now contains `A2`, and the
whole document is written. That is already the behaviour, it is correct, and inferring "this looks
like a rename" would be exactly the second-guessing this rule exists to prevent.

### The live path has to emulate it

The object model has no whole-document write, so `ALL` on the live path is necessarily update-every-
query, add-missing, remove-extra. **The requirement is equivalence of outcome, not of mechanism**: a
workbook synced live under `ALL` must end up holding what it would have held had the file been
closed. That is testable the same way the round trip in `PQ-15` is testable, and it should be tested,
because it is the one place where "the same command did two different things" could come back.

## A manifest that disagrees with the document is an ERROR

Not a guess. If the manifest names a query the document does not define, or the document defines a
query the manifest does not name, **stop and say so**.

The earlier proposal - treat a name in the manifest but absent from the document as a deliberate
deletion - was inference dressed up as a rule. Both readings are plausible (deliberate deletion, or
stale manifest), the cost of choosing wrong is deleting somebody's query, and there is a third option
that costs nothing: refuse, name the discrepancy, and let a human resolve it.

```
Manifest and document disagree.
  In the manifest but not defined:  RawInput
  Defined but not in the manifest:  FinalTable
Fix the manifest or the document, then sync.
```

This costs the hand-roller one accurate list, which is a fair price for never guessing at intent, and
it makes the manifest self-checking: a typo is caught the first time rather than the day it deletes
something.

**Offering to fix it is not guessing** and would be a kindness - "the document defines `FinalTable`,
which the manifest does not name. Add it?" - as long as it asks. Worth building only after the error
exists and people complain about it, which is the honest way to find out whether they mind.

## The unresolved tension: confirming removals under `ALL`

Confirming removals and leaving `ALL` untouched cannot both be fully true, and this needs a decision
rather than a compromise invented in code.

Today, `ALL` on the file path removes queries **silently**, because it does not diff and therefore
does not know what is being removed. Adding a confirmation means diffing - which is affordable, since
the existing DataMashup is already parsed - but it introduces a prompt where there has never been
one, for thousands of people, on a workflow they run daily.

Three options, in order of preference:

1. **Confirmation defaults ON for what is new, OFF for what is not.** A list manifest and live sync
   are both new, so they may prompt from day one. `ALL` on the file path keeps its current silence
   unless the user opts in. Nothing anybody relies on changes, and the risky new paths are guarded.
2. **On everywhere, with a setting to silence it.** Safer for the stale-`ALL` case, but it changes a
   daily workflow for every existing user and will be experienced as the extension becoming naggy.
3. **Off everywhere by default.** Consistent and cheap, and gives up the protection that made the
   confirmation worth having.

Recommending (1). It respects the rule that `ALL` behaves as it always has, while not shipping a new
delete-capable path with no brakes.

## Slices

- `PQ-22` — the `Queries:` manifest, and reading it. **Gates the rest.**
- `PQ-23` — honour it on both paths, with `ALL` on the file path left exactly as it is
- `PQ-24` — "Extract Selected Power Queries", writing the manifest
- `PQ-25` — confirm removals; see the tension above before choosing a default
- `PQ-26` — never write `ALL` after a partial extraction
- `PQ-27` — error, do not guess, when manifest and document disagree

Nothing here changes what `ALL` does. If a slice starts to, it is the slice that is wrong.
