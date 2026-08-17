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

**v0.6.0 = live sync to an open workbook.** The mechanism is proven (`PQ-12`); the rest is product.

**Live sync works on a real workbook.** What is left is getting it to the person who asked.

The pipeline works and cannot publish by accident. What is left is a decision, not a blocker.

1. `PQ-19` — CI has failed on every run for a year; macOS cannot launch VS Code
2. `PQ-16` — reply to namgaw, who asked for this in October 2025
3. Merge `wip/live-sync`, decide on 0.6.0, and set `MARKETPLACE_PUBLISH` when ready
2. `PQ-01` — the pipeline; nothing can be released until it works
3. `PQ-10` — one README, before the marketplace page is republished

## Last Numbers

Read this **before** minting an ID. It sits above the tables because it is an input to writing one,
not a summary of them.

| Series | Last Num | Series Description |
| --- | --- | --- |
| PQ | PQ-31 | Excel Power Query Editor slices and fixes |

`Last Num` is a cache over the Delivery Index, not a second source of truth — the IDs in the tables are
authoritative. When minting, take the next number **and** confirm it is unused across every sub-table,
then bump this cell in the same edit.

## Delivery Index

**Rows are one line.** `Doc` pins a filename; anything wanting a paragraph wants a slice doc.

States: `⬜ planned` · `⛔ blocked` · `🟨 coded` — built and deployed, nobody has used it yet · `💨 proven` — someone
used it and it worked · `⏸ retired` — tried, backed out, kept for the reason.

### Shipping — the pipeline is the blocker

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-01 | 💨 proven | Rebuild release.yml as tag-triggered | M | — | proven 2026-08-15 — v0.6.0-rc.1 built a draft, marketplace skipped |
| PQ-02 | ⬜ planned | Ship the settings refactor as 0.6.0 | S | — | unblocked; NOT a patch — 13 settings renamed |
| PQ-03 | ⬜ planned | Verify marketplace publish end to end | S | — | `VSCE_PAT` path has never demonstrably fired from a tag |
| PQ-04 | 💨 proven | Prune the release workflow | S | — | done with PQ-01 — 318 lines to 190 |

### v0.6.0 — live sync to an open workbook

The headline feature, requested by a real user in October 2025 and left for ten months because the
object model looked like a dead end. It is not: `WorkbookQuery.Formula` is read/write, and a running
Excel serves external automation. See `design/live-sync-to-open-excel.md`.

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-12 | 💨 proven | Spike: rewrite a query in a workbook the user has OPEN | S | live-sync-to-open-excel.md | proven 2026-08-14 — attached via ROT, wrote, went dirty |
| PQ-13 | 💨 proven | Helper process, status/write, error paths | L | live-sync-to-open-excel.md | proven 2026-08-14 — ROT lookup, retry on busy, stdin payload |
| PQ-14 | 💨 proven | Split a section document and match queries by name | M | live-sync-to-open-excel.md | proven 2026-08-14 — 3 fixtures round-trip byte for byte |
| PQ-15 | 💨 proven | Round-trip test: section -> N formulas -> Excel -> section | M | live-sync-to-open-excel.md | proven 2026-08-14 — byte-identical through a real Excel |
| PQ-16 | ✅ done | Reply to namgaw, and reach out to Ken Puls | S | [discussion #3](https://github.com/ewc3labs/excel-power-query-editor/discussions/3) | replied with the built feature and a prerelease he can install; his Monkey Tools pointer is what unstuck it. Ken Puls NOT contacted, deliberately - he is credited in the README, which is better than cold-emailing the author of the commercial tool we just built a free alternative to |
| PQ-17 | 💨 proven | Wire live sync into the sync command and settings | M | live-sync-to-open-excel.md | proven 2026-08-14 — real 29-query workbook in a OneDrive folder |

### Selective extract, and who is authoritative

The two write paths currently disagree about deletion, and nobody chose that. See
`design/selective-extract-and-sync-authority.md`.

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-22 | ⬜ planned | `Queries:` manifest in the .m header, ALL or a block list | M | selective-extract-and-sync-authority.md | gates the rest; absent means ALL, matching every existing file |
| PQ-23 | ⬜ planned | Honor the manifest on both paths; `ALL` untouched | L | selective-extract-and-sync-authority.md | file+ALL stays setFormula(document); live must MATCH its outcome |
| PQ-24 | ⬜ planned | "Extract Selected" command, writing the manifest | M | selective-extract-and-sync-authority.md | second menu item, so extracting all costs nobody a decision |
| PQ-26 | ⬜ planned | Never write `Queries: ALL` after a partial extraction | S | selective-extract-and-sync-authority.md | a document that lies about being complete deletes what it could not read |
| PQ-25 | ⬜ planned | `sync.confirmQueryRemoval`, defaulted ON | S | selective-extract-and-sync-authority.md | changes `ALL` behavior on purpose — changelog must say so |
| PQ-27 | ⬜ planned | Manifest vs document mismatch is an error, not a guess | S | selective-extract-and-sync-authority.md | name the discrepancy, refuse, let a human resolve it |

### Docs

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-28 | ✅ done | CONFIGURATION.md still documents 13 renamed settings | S | [Config_Reference](../Config_Reference.md) | replaced by a reference generated from package.json; renames moved to Config_Changes |
| PQ-29 | ✅ done | USER_GUIDE.md has no mention of live sync | S | [Live_Sync](../Live_Sync.md) | feature doc written; User_Guide now carries a section pointing into it |
| PQ-30 | ✅ done | Retire the RELEASE_SUMMARY pattern | S | — | removed; CHANGELOG and the release body are the two homes |
| PQ-32 | ✅ done | docs: link and orphan checking in CI | S | [Overview](../Overview.md) | `npm run docs:links` found 11 dead links and 1 unreachable doc on its first run |
| PQ-34 | ⬜ planned | Marketplace pre-release channel, and the semver it requires | M | [PQ-34](slices/PQ-34_Marketplace_Prerelease_Channel.md) | the button is easy; the cost is a permanent odd/even minor convention, because publishing 0.6.0 as pre-release burns it for stable. Deferred until there is more than one tester |
| PQ-31 | 🟡 partial | bump-version: drop commit analysis, sync the README badge | S | [PUBLISHING_GUIDE](../PUBLISHING_GUIDE.md) | badge sync DONE by docs-tools `values`; commit analysis still there, and the `npm version` tag hazard is now documented rather than fixed |

### Data safety — the thing that must never break

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-33 | 🟡 measured | AutoSave vs live sync: untested interaction, and backups per sync | M | [PQ-33](slices/PQ-33_AutoSave_And_Live_Sync.md) | MEASURED: AutoSave commits a live write in ~2s and closing without saving does NOT undo it. Docs corrected, message differentiated. Backup churn was already answered by retention |
| PQ-05 | ⬜ planned | Audit every path that writes a workbook | M | — | backup-then-temp-then-swap, no exceptions |
| PQ-06 | ⬜ planned | Prove `.xlsb` round-trips byte-for-byte | M | — | binary and unforgiving; fixtures exist |

### Structure

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| PQ-07 | ⬜ planned | Extract the workbook read/write seam from extension.ts | L | — | 2,152 lines in one file; extract WITH tests, never wholesale |
| PQ-08 | ⬜ planned | Settings deprecation policy | S | — | 18 public settings; renaming one breaks configs silently |
| PQ-09 | 💨 proven | Replace the settings WIPE with a real migration | M | settings-migration.md | proven 2026-08-14 — 7 legacy values migrated on a real install |
| PQ-10 | 🟨 coded | Converge on ONE README | S | — | merged, split files and swap scripts deleted; needs a read |
| PQ-11 | 💨 proven | Research: does the PQ/M extension ship Excel symbols | S | excel-symbols.md | answered 2026-08-14 — it does NOT ship Excel.CurrentWorkbook |
| PQ-18 | 💨 proven | Push symbols through the Power Query API | M | excel-symbols.md | proven 2026-08-15 — survives the PQ extension being absent or added later |
| PQ-19 | 🟨 coded | CI: get a real signal out of the matrix | M | — | fail-fast off + test timeout; windows-22 GREEN, first in a year |
| PQ-20 | ✅ done | CI: macOS cannot launch VS Code at all | M | — | TWO bugs behind one excuse: test-electron 2.5.2 spawned `Contents/MacOS/Electron`, renamed to `Code` in VS Code 1.110+; then a 106-char user-data socket path against macOS's 104-byte limit. Both fixed, macOS green, continue-on-error removed |
| PQ-21 | 🟨 coded | CI: ubuntu fails, and PQ-18 was part of why | M | — | workspace error GONE from CI; both windows legs now green |

## Working Rules

- A slice earns a row once the problem is understood **with receipts**, not when it feels important.
- Rows stay one line. If it needs a paragraph, it needs a doc in `slices/`.
- Reconciling the punchlist means asking *"have we already got a slice for this?"* before minting a new
  one — the whole point is not recording the same thing three times.

## Related Documentation

- `EPQE_Checkin_Punchlist.md` — things noticed, not yet understood
- `../../AGENTS.md` — how to work in this repo, and what must never break
- `../RAG_Sessions/` — how a hard problem actually got solved
