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
test for `\\` catches only paths that were never abbreviated to a drive letter. So the failure
decides, in `src/resilientWatcher.ts`:

| state | on error |
| --- | --- |
| native | build a **polling** watcher, then close the dead one — no moment with nothing watching |
| polling | **stop**, close it, and report the file unwatchable — never loop |
| stopped | ignore — including errors that arrive after the user stopped watching |

Errors from a watcher that has already been **replaced or closed are ignored**, and the transition
is claimed before any `await`, so two errors in one tick build one fallback, not two.

When a file becomes unwatchable, it is **removed from the watch registry and the status bar**,
rather than staying listed as watched. In a dev container the VS Code backup watcher is still
running, so the file stays watched and only the log says chokidar failed.

Polling is slower and works everywhere, which is the right trade for a path that has just proved it
cannot do better.

### The first version, and what review changed

It lived inside `extension.ts`, and this slice said it could not be tested because *"faking the
error would test the fake."* **That was wrong**, and the repository's own `AGENTS.md` says so: new
functionality gets its own module, with tests. The thing needing tests is the **state machine**, and
an injected watcher tests exactly that. It also had two real defects:

- **A second failure left the file registered**, so the status bar and **Toggle Watch** kept
  reporting a watcher that would never fire. *Found by review.*
- **Stopping watching did not stop the fallback.** Teardown closed the raw chokidar watcher, so a
  late error from it could start a polling watcher nothing held and nothing would close. The
  registry now holds the resilient wrapper, and closing it moves the state to `stopped`. *Found
  while extracting, not by review.*

## Tests

`test/resilientWatcher.test.ts`, eight cases against an injected watcher: it starts native with
handlers attached; native→polling swaps once and closes the dead watcher; two errors in one tick
build one fallback; a late error from the replaced watcher is ignored; a polling failure stops
instead of looping; starting in polling mode has nothing to fall back to; an error after close does
nothing; close after a fallback reaches the live watcher.

**Not tested, and only a real share can test it:** whether chokidar's polling actually detects saves
on a network drive. That is why this stays 🟨.

## To prove

On the work PC, watch a `.m` file on `P:\`. The log should show the native error, then
`Retrying with polling`, then `Watcher ready ... (polling: true)` — and a save should trigger a
sync.

Related: [PQ-35][pq-35] — the same drive also broke the live sync lookup.

[pq-35]: PQ-35_Live_Sync_On_Mapped_Network_Drives.md
