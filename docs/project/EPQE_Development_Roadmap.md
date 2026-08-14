<div align="center">
<table>
<tr>
<td style="padding: 0 40px; vertical-align: middle;">
<img src="../../assets/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="72" height="72">
</td>
<td style="vertical-align: middle;">
<h1 style="margin: 0;">Excel Power Query Editor</h1>
<h3 style="margin: 5px 0;"><strong>Development Roadmap — where we are, and what is next</strong></h3>
</td>
</tr>
</table>
</div>

---

AsOf: 2026-08-14

## Current Focus

The extension works and has users. What is broken is everything *around* it: a release pipeline that
last succeeded in July 2025, and a finished version that never shipped because of it.

1. `PQ-01` — next, and it unblocks everything else
2. `PQ-02` — the release that has been sitting built for a year
3. `PQ-05` — before any refactor touches the write path

## Last Numbers

Read this **before** minting an ID. It sits above the tables because it is an input to writing one,
not a summary of them.

| Series | Last Num | Series Description |
| --- | --- | --- |
| PQ | PQ-08 | Excel Power Query Editor slices and fixes |

`Last Num` is a cache over the Delivery Index, not a second source of truth — the IDs in the tables are
authoritative. When minting, take the next number **and** confirm it is unused across every sub-table,
then bump this cell in the same edit.

## Delivery Index

**Rows are one line.** `Doc` pins a filename; anything wanting a paragraph wants a slice doc.

States: `⬜ planned` · `🟨 coded` — built and deployed, nobody has used it yet · `💨 proven` — someone
used it and it worked · `⏸ retired` — tried, backed out, kept for the reason.

### Shipping — the pipeline is the blocker

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-01 | ⬜ planned | Rebuild release.yml as tag-triggered | M | — | `workflow_run` runs as main; tag branches unreachable |
| PQ-02 | ⬜ planned | Ship 0.5.2 | S | — | built ~2025-07; package.json is ahead of the newest release |
| PQ-03 | ⬜ planned | Verify marketplace publish end to end | S | — | `VSCE_PAT` path has never demonstrably fired from a tag |
| PQ-04 | ⬜ planned | Prune the release workflow | S | — | 318 lines, mostly reporting; RecallTape's is a third of it |

### Data safety — the thing that must never break

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-05 | ⬜ planned | Audit every path that writes a workbook | M | — | backup-then-temp-then-swap, no exceptions |
| PQ-06 | ⬜ planned | Prove `.xlsb` round-trips byte-for-byte | M | — | binary and unforgiving; fixtures exist |

### Structure

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-07 | ⬜ planned | Extract the workbook read/write seam from extension.ts | L | — | 2,152 lines in one file; extract WITH tests, never wholesale |
| PQ-08 | ⬜ planned | Settings deprecation policy | S | — | 18 public settings; renaming one breaks configs silently |

## Working Rules

- A slice earns a row once the problem is understood **with receipts**, not when it feels important.
- Rows stay one line. If it needs a paragraph, it needs a doc in `slices/`.
- Reconciling the punchlist means asking *"have we already got a slice for this?"* before minting a new
  one — the whole point is not recording the same thing three times.

## Related Documentation

- `EPQE_Checkin_Punchlist.md` — things noticed, not yet understood
- `../../AGENTS.md` — how to work in this repo, and what must never break
- `../RAG_Sessions/` — how a hard problem actually got solved
