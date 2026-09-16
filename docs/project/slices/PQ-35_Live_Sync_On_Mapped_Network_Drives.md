# PQ-35 — Live sync on mapped network drives

**State:** 🟨 coded · **Est:** S · Minted 2026-09-14 · Reported by Wilson from the work PC

## The problem

Live sync could not reach a workbook open from a mapped network drive, and **the message blamed the
wrong thing.**

```text
[syncToExcel] Excel file is locked; live sync cannot handle it (available=true, excelProcesses=1)
[syncToExcel] Excel is running but this workbook is not visible to the extension. This usually means
              one of them is elevated and the other is not - COM hides running objects across
              integrity levels.
```

Neither VS Code nor the helper was elevated. The workbook was open, as that user, in the one running
Excel.

## What was actually happening

The helper finds an open workbook by walking the Running Object Table and binding only to a display
name that is already registered — it never opens anything. It compared the path the user sees
against what COM registered, **and those are two different strings for the same file:**

| asked | answer |
| --- | --- |
| path VS Code has | `P:\IT\MedAR_Coder\MedARCoder_OpenRecords.xlsx` |
| Excel's `Workbook.FullName` | `P:\IT\MedAR_Coder\MedARCoder_OpenRecords.xlsx` |
| **Running Object Table display name** | **`\\medarms01\public\IT\MedAR_Coder\MedARCoder_OpenRecords.xlsx`** |

Excel registers a network workbook under its **UNC path** even when it was opened through the drive
letter. The exact-path lookup found nothing and fell through to "not open".

### The measurement that nearly hid it

Checking `Workbook.FullName` returned `P:\...` and appeared to rule the UNC theory out. **It did
not.** `FullName` is what Excel *reports*; the ROT display name is what COM *registered*, and only
the second is what the helper matches against. The theory was confirmed only by dumping the table
itself with the helper's own `RunningObjects.Names()`.

A second false lead cost a round trip: the first dump was taken with the workbook closed, which
looked like "Excel never registered it". Reopened, it was there under the UNC name.

### Why it took three round trips at all

The helper's not-found response stated a conclusion — `open: false` — and **withheld the evidence.**
The workbook was registered the whole time, under a name we did not try. Diagnosing that took
hand-pasted COM code, including one command (`Marshal.GetActiveObject`) that does not exist in
PowerShell 7 — a trap already documented in `excelLive.ts`.

## The fix

1. **Try the UNC form.** `NetworkPaths.ToUnc` asks Windows what the drive letter maps to
   (`WNetGetConnection`) and the helper tries that exact name after the exact path and before cloud
   URLs. This is an **identity transform, not a heuristic** — it is the same mapping Excel resolved
   — so it stays inside the helper's no-scoring rule. Local, SUBST, already-UNC and unparseable
   paths return null.
2. **Report what the helper could see.** When nothing matches, the response now includes
   `registered`: every workbook display name in the table. The extension logs them. **If the table
   cannot be read, `registered` is absent, not empty.** An empty list is evidence of an integrity
   wall, and a failed read is no evidence. The first version conflated them in both the helper and
   `RunningObjects.Names()`.
3. **Let the evidence choose the message.** `explainInvisibleWorkbook` used to say "usually
   elevated" unconditionally, including for a helper that had just measured itself not elevated.

   | the helper saw | message |
   | --- | --- |
   | itself elevated | VS Code is running as administrator |
   | Excel running, **zero** workbooks visible | an integrity wall — Excel is probably elevated |
   | Excel running, **some** workbooks visible, not this one | a different path, or not open — names in the log |
   | an older helper, no evidence | both possibilities, ranked neither way |

This is the third time a live-sync message stated a cause the evidence already contradicted — see
`FIX-5`. The pattern is the defect: **a message should be derived from what was measured.**

## Tests

- `Why a running Excel cannot see the workbook` — every message branch, including that visible
  workbooks rule elevation out.
- `Mapped network drives` — the lookup order is exact, then UNC, then cloud; and the shipped C#
  still **compiles**, since a compile error in `RunningObjects.cs.txt` breaks every live sync, not
  only network ones. `ToUnc` declines local, UNC, empty and relative paths.

**Not covered by CI, and cannot be:** the positive case needs a real mapped drive, which a runner
cannot create without changing the machine.

**What was verified by hand, and what was not.** On 2026-09-14, a ROT dump showed Excel registering
a `P:\` workbook under `\\medarms01\public`. That is the *diagnosis*. **`ToUnc` itself has never run
against a mapped drive, and live sync has never succeeded end to end on one.** The first version of
this slice said the positive case was "verified by hand", which a review correctly flagged as
contradicting the state and the section below.

## To prove

Install the build on the work PC, open a workbook from `P:\`, and live-sync to it. The log should
read:

```text
[syncToExcel] Excel file is locked; live sync CAN handle it (available=true, found via exact-unc as \\medarms01\public\...)
```

**This step was originally unobservable.** It named `matchedHow`, a field the helper always returned
and the extension never parsed or printed, so the only available proof was inferring the UNC match
from the sync succeeding. Caught while preparing the work-PC test build; the log line now prints it.

## Known limits

- **DFS namespaces.** A drive mapped to `\\domain\dfs\...` may be registered under the namespace or
  the resolved target. `ToUnc` returns what the drive was mapped to; if Excel registers the other
  form, the `registered` list in the log will show it.
- The reverse — VS Code opened via UNC while Excel registered a drive letter — was not observed and
  is not handled. Excel registered UNC here, so an exact match already covers a UNC-opened `.m`
  file.

Related: [PQ-36][pq-36] — the same drive broke the watcher too.

[pq-36]: PQ-36_File_Watcher_On_Network_Drives.md
