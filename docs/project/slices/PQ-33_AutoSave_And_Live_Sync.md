# PQ-33 — AutoSave vs live sync

**State:** 🟡 measured · **Est:** M · Minted 2026-08-16 · Measured 2026-08-16

## The problem

Live sync is designed around one promise: **the file on disk is not touched, and the workbook is
left dirty so you can review the change and decide whether to save it.** That promise is what makes
it safe to write into a workbook somebody is using.

**AutoSave may remove that promise entirely, and we have never tested it.**

AutoSave is on by default for workbooks stored in OneDrive and SharePoint — which is exactly the
population live sync was built for, and exactly what was used to prove the feature works. With
AutoSave on, Excel persists changes continuously. So a live write is plausibly saved to the cloud
within seconds, with no review step, no chance to close-without-saving, and a version now in the
file history.

None of that is necessarily wrong. It may be fine, or even what the user wants. But it is currently
**unknown**, and the documentation states the opposite as a guarantee.

## The second half: a backup per sync — mostly already answered

`syncToExcel` copies the workbook before every sync, so a fast edit-sync loop produces a backup per
save.

**This is largely by design and already handled.** `backup.maxFiles` exists precisely because that
churn showed up in practice; retention was the fix, and it works. Debouncing is therefore a
refinement rather than a defect.

What is left is narrower: with a low `maxFiles` and a rapid session, the state you actually wanted
can be pushed out of the retention window by later backups of a worse state. Worth considering, not
worth blocking anything over.

For cloud workbooks there is also a second safety net — OneDrive and SharePoint version history
recovers the whole workbook, which makes turning backups off a defensible choice there rather than a
reckless one. The documentation now says so.

## Measured 2026-08-16

Two scratch copies of `test/fixtures/simple.xlsx`, one in a OneDrive synced folder and one local, in
a dedicated Excel instance. A query formula was set through COM - the same call live sync makes -
and then `Saved`, `AutoSaveOn` and the file's mtime were watched.

| | after the write | ~2s | ~5s | closed WITHOUT saving |
| --- | --- | --- | --- | --- |
| **Cloud**, `AutoSaveOn=True` | `Saved=False` | `Saved=True` | disk mtime changed | **change survived** |
| **Local**, `AutoSaveOn=False` | `Saved=False` | `Saved=False` | unchanged | change discarded |

**Answers:**

1. **Yes, and within about two seconds.** AutoSave commits a live write with no user action.
2. **`Saved` returns to `True` almost immediately**, so `sync.requireSavedWorkbook` will rarely fire
   on a cloud workbook. It is not useless: the risk it guards against - unsaved work destroyed by
   our write - barely exists there either, because AutoSave has already put those edits on disk
   where the backup can capture them. **The guard protects local workbooks**, which is where the
   danger is.
3. **The review step does not exist for AutoSave workbooks.** Closing without saving does not undo
   it. Version history is the undo path.
4. **`Workbook.AutoSaveOn` is readable through COM** and was `True` on the real 29-query workbook
   live sync was proven against.

## Done as a result

- The helper reports `autoSaveOn` on both status and write responses.
- The completion message tells the truth per workbook: AutoSave workbooks are told Excel saves it
  itself and that version history is the undo path, rather than being promised a review step.
- `Live_Sync.md` carries the measurement, and the headline "close without saving and the edit is
  gone" promise is now qualified rather than stated flatly.

## Still open

- Whether to warn BEFORE writing to an AutoSave workbook rather than explaining afterwards. Leaning
  no: it is not dangerous, version history covers it, and a confirmation on every sync is the kind
  of friction that gets a feature turned off.
- Debouncing the per-sync backup, which retention already largely handles.

## What was originally to be established

Measured, not reasoned about:

1. With AutoSave **on**, does a live write get persisted automatically? How quickly?
2. Does `Workbook.Saved` ever report `false` for an AutoSave workbook, or is it effectively always
   `true`? If the latter, `sync.requireSavedWorkbook` never fires for cloud workbooks and the
   protection it offers is illusory there.
3. Does the AutoSave version history capture our write as its own version — i.e. is there a real
   undo path that makes the missing review step acceptable?
4. Can AutoSave be detected through COM (`Workbook.AutoSaveOn`) reliably enough to change behavior
   or at least to say so in the log?

## Likely outcomes

- **Detect and disclose.** If `AutoSaveOn` is readable, say in the log and in the docs what will
  actually happen, rather than promising a review step that does not exist.
- **Possibly debounce the backup.** One per workbook per session rather than per sync. A refinement
  now that retention is understood to be the real answer to churn.
- **Possibly nothing else.** If version history genuinely covers the user, the honest fix is to
  correct the documentation and move on.

## Why it is not a blocker

Live sync is off by default and labeled beta. Anyone who has turned it on has opted into something
explicitly experimental. This needs to be settled before live sync is described as ready, not before
it is available.

## Origin

Raised by Wilson on 2026-08-16, while adding `sync.requireSavedWorkbook`: the observation that
AutoSave "syncs back to Excel every sync, and we have not yet invented a way to debounce this backup
on live sync."
