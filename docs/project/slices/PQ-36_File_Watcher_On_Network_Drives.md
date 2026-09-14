# PQ-36 — The file watcher is deaf on network drives

**State:** 🟨 coded · **Est:** S · Minted 2026-09-14 · Reported by Wilson from the work PC

## The problem

On a mapped network drive, watching a `.m` file reported success and then never fired:

```text
13:28:08.562 [watch] Started watching: MedARCoder_OpenRecords.xlsx_PowerQuery.m
13:28:08.573 [watchFile] CHOKIDAR: Watcher ready for MedARCoder_OpenRecords.xlsx_PowerQuery.m
13:28:08.575 [watchFile] CHOKIDAR: Watcher error: Error: UNKNOWN: unknown error, watch
```

**Ready, then an error 2ms later, then silence.** Every subsequent save was unwatched, and nothing
told the user. A watcher that says `ready` and is deaf.

## Why

Chokidar uses native `fs.watch` unless told to poll, and `fs.watch` depends on local change
notifications that SMB shares do not reliably provide. The watcher only polled in dev containers:

```ts
usePolling: isDevContainer
```

The `error` handler logged and did nothing else.

## The fix

**Detect, do not predict.** Node cannot tell whether `P:\` is a local disk or a mapped share, and a
test for `\\` catches only paths that were never abbreviated to a drive letter. So rather than
classifying the drive up front, the error handler now:

1. on the first native failure, closes the watcher and **retries with polling** (1s interval),
   logging that this usually means a network drive or UNC path;
2. **repoints the registered watcher set** at the new instance — otherwise teardown closes the dead
   watcher and leaks the live one;
3. if polling fails too, says so plainly in the log and a warning, and points at manual **Sync to
   Excel** rather than retrying forever.

Polling is slower and works everywhere, which is the right trade for a path that has just proved it
cannot do better.

## Tests

Type-check and lint pass. **No automated test**: reproducing `fs.watch` failing needs a real network
share, and faking the error would test the fake. The recovery path is small and linear.

## To prove

On the work PC, watch a `.m` file on `P:\`. The log should show the native error, then
`Retrying with polling`, then `Watcher ready ... (polling: true)` — and a save should trigger a
sync.

Related: [PQ-35][pq-35] — the same drive also broke the live sync lookup.

[pq-35]: PQ-35_Live_Sync_On_Mapped_Network_Drives.md
