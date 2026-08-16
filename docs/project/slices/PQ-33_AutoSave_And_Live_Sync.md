# PQ-33 — AutoSave vs live sync

**State:** ⬜ planned · **Est:** M · Minted 2026-08-16

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

## What has to be established

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
