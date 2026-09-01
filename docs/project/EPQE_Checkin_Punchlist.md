<div align="center">
<table>
<tr>
<td style="padding: 0 40px; vertical-align: middle;">
<img src="../../assets/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="72" height="72">
</td>
<td style="vertical-align: middle;">
<h1 style="margin: 0;">Excel Power Query Editor</h1>
<h3 style="margin: 5px 0;"><strong>Check-in Punchlist — stuff we noticed, and what is still open</strong></h3>
</td>
</tr>
</table>
</div>

---

Started: 2026-08-14

## How this works

Things we saw and need to address **sometime**. No evidence required to write one down — that is the
point. An item can sit here indefinitely without guilt.

An item becomes a roadmap slice once someone has done enough analysis to **document the problem with
receipts**. Before minting a new slice, check whether an existing one already covers it.

**Conventions that span every EWC3 Labs repo are tracked separately**, so this list stays about this
extension. If an item here turns out to be about the whole estate, it moves and the entry goes away.

---

## 2026-09-01 — a Windows CI flake, caught by yesterday's guard

**Items:**

- [x] **`Watch Tests` teardown failed on windows-latest / node 24 with `EPERM` from `rmSync`.** A
      file watcher still held a handle on the temp directory, so Windows refused to remove it and an
      `after all` hook turned a green run red. The same commit had passed minutes earlier, which is
      what identified it as a race rather than a defect.

      Fixed with `maxRetries`/`retryDelay` — Node's documented answer for EBUSY/EPERM on Windows —
      and a `try/catch` so cleanup can never fail a suite again. `force: true` was never going to
      help; it only ignores ENOENT. Applied to all four temp-directory teardowns, since they had
      identical exposure and only one had been unlucky.

- [x] **`FIX-4` earned itself in production the day after it was written.** The failing run printed
      *"Tests failed (exit 1); test-counts.json left unchanged. The partial run counted 143."*
      Before that guard, a partial run would have written 143 over 142, and the next green build
      would have failed on a count mismatch instead of the real problem.

## 2026-09-01 — refresh after sync, and what Excel is doing while we write
**Items:**

- [ ] **Live sync writes the query and never refreshes it, and nothing says so.** The helper sets
      `Workbook.Queries(name).Formula` and stops; the loaded table keeps showing the previous result
      until somebody hits **Data → Refresh All**. That is defensible - a refresh can be slow, can
      hit a data source, and can prompt - but right now a successful sync looks like nothing
      happened, which is the worst version of the choice. At minimum say so in the success
      notification.

- [ ] **Wanted: refresh on sync, off by default, with a level rather than a switch.** Sketch: `off`
      (today), `updated` - refresh only the queries whose M actually changed, which we already know
      because live sync diffs per query - and `all`, mapping to `Workbook.RefreshAll()`. `updated`
      is the interesting one and the reason a plain checkbox is wrong: `RefreshAll` fires every
      connection in the workbook, including ones this extension never touched.

      Risks to settle **before** building it, not after:
      - a refresh can raise a **modal prompt** - privacy levels, credentials, disabled external data
        - and a modal in an app we are automating is exactly what our retry logic cannot outwait
      - `BackgroundQuery = true` makes Refresh return immediately and surface failures nowhere, so
        "success" would mean "we asked", not "it worked"
      - long refreshes against real sources blow past the helper timeout
      - a refresh that fails halfway leaves the workbook in a state we did not choose and cannot
        describe

      **Must be tear-outable.** If it destabilises anything, deleting the refresh step has to leave
      the write path exactly as it is today. No shared state, no reordering, no "while we are here".

- [ ] **UNTESTED, AND THE ONE THAT WORRIES ME: what happens when the Power Query editor or the Get
      Data wizard is open while we write?** The editor is modal to Excel, so the likely outcome is
      that our busy-retry - 150/300/450/600ms, written on the assumption that *"Excel dialogs are
      brief"* - exhausts and the write fails before touching anything. That is the safe result, and
      the message the user gets would still be a raw COM error rather than "close the Power Query
      editor and try again".

      The scenario that actually needs proving is the other one: **if the write succeeds while the
      editor holds that query open, does Close & Load overwrite it?** The sync would report success
      and the user's edit would silently revert. Not corruption, but indistinguishable from it in
      the moment.

      Test both before shipping refresh-on-sync, because refresh makes Excel busier at exactly the
      moment we are least sure what it is doing.

- [ ] **Persistent busy deserves its own message.** We already retry `RPC_E_CALL_REJECTED` and
      friends. When those retries exhaust, we know something modal is open - which is worth saying
      plainly rather than reporting an HRESULT. Cheap, and it covers the PQ-editor case without
      needing to detect the editor specifically.
## 2026-08-14 — first pass, adopting the documentation conventions
**Items:**

- [ ] **The two write paths disagree about deletion and nothing says so.** File sync replaces the
      whole section document, so a query absent from the .m is deleted. Live sync diffs per query
      and never deletes. Same command, same file, opposite outcome depending on whether Excel
      happens to be open. Neither behavior is documented and neither was chosen. [PQ-23]

- [ ] **The dev loop produces false negatives.** Reinstalling the extension does not affect a
      RUNNING VS Code: the extension host keeps executing whatever it loaded at startup, so a fix
      can be built, packaged, installed and verified on disk while the editor still runs the old
      code. That cost a full debugging cycle on live sync - the code was correct and the symptom
      said otherwise. Worse, `package.json` stays at the same version through dozens of dev
      installs, and VS Code keys the extension folder by version, so same-version reinstalls are not
      reliably a clean replace either. `npm run bump-version` exists and is not being used in the
      loop. Wanted: a dev-install script that bumps a prerelease number, packages, installs, and
      either reloads the window or says plainly that a reload is required.
- [ ] **And it must install into the editor that is actually RUNNING.** `code` and `code-insiders`
      are separate installs with separate extension folders. A whole debugging session went into
      live sync "not working" while every build was landing in stable VS Code and the developer was
      running Insiders, which still had 0.5.1 and no live-sync helper at all. Four real bugs were
      found and fixed underneath that, each of which looked like the cause and none of which was.
      The dev-install script should detect the running variant (or install to both) and print which
      folder it wrote to.

- [x] **What is 0.5.2?** A settings refactor. Every setting was moved into a namespace:
      `watchAlways` -> `watch.always`, `logLevel` -> `log.level`, `syncTimeout` -> `sync.timeout`,
      and so on — **13 of 19 renamed or removed**, with `debugMode` and `verboseMode` gone entirely.
      There is no migration code. Shipping it as a patch would silently reset the configuration of
      every existing user. [PQ-09] blocks [PQ-02].
- [x] **`migrateLegacySettings()` is a WIPE, not a migration**, and it runs at activation
      (`extension.ts:306`). It enumerates the config and sets every key to `undefined` in both User
      and Workspace scope, preserving nothing, and its guard compares against the *extension
      version* so it re-fires on every bump. The correct earlier attempt sits commented out directly
      above it. Full write-up in `docs/analysis/settings-migration.md`. [PQ-09]
- [x] **Why removing old settings was impossible:** VS Code only lets an extension write a key that
      is *registered*. Once the old names were deleted from `contributes.configuration`, they could
      no longer be cleared — the value sits in the user's settings.json permanently as an unknown
      setting. The fix is to keep them declared with `deprecationMessage`, migrate, then clear.
      [PQ-09]
- [x] ~~The code still references the OLD setting names.~~ **Audited 2026-09-01: no longer true.**
      Every remaining occurrence is legitimate — local variable names that read the *new* keys
      (`const logLevel = config.get('log.level')`), the migration map which must name old keys by
      definition, and the migration function that deliberately reads `debugMode`/`verboseMode` in
      order to clear them. Not one is a read expecting an old value. [PQ-09]
- [x] **The README swap:** `docs/README.gh.md` and `docs/README.vsmarketplace.md` were both **0
      bytes**, emptied in `a2ea1ef` ("save before settings refactor") — a work-in-progress commit.
      Recovered from `fb52fc0` (v0.5.0): 8,034 and 4,034 bytes respectively.
- [x] ~~`scripts/set-readme-gh.js` and `set-readme-vsce.js`.~~ **Decided: one README.** The swap
      scripts and both source files are gone; `README.md` serves GitHub, the Marketplace and the VS
      Code extension pane, and `vsce` rewrites relative image paths to `raw/HEAD` on the way out.
      One document to keep honest instead of three to keep synchronised. [PQ-10]
- [x] ~~EPQE writes `excel-pq-symbols.json` into `.vscode/`.~~ **Fixed by PQ-18: it writes nothing
      at all now.** Symbols are handed to the Power Query extension through `addLibrarySymbols`
      (`vscode-powerquery#206`), so nothing lands in the user's folder, no setting is touched, and
      the plain-folder case stops mattering. `FIX-2` reports leftovers from the old file-based
      version without deleting them — confirmed firing in the wild in a user's log. [PQ-11]
- [ ] **Build real test fixtures.** `test/fixtures/test_workbook.xlsx` and `.xlsb` are not workbooks
      — the first contains the literal text "test xlsx file", and `test_clean.xlsx` is four bytes.
      Byte-identical to v0.5.0, so they have always been stubs, and any test that "opens" one proves
      nothing. **Decided 2026-09-01: make them for real.** Genuine workbooks with genuine queries,
      copied into a work folder per test and mutated there, so a test can extract, edit and sync
      against something that behaves like a user's file. Wanted: `.xlsx`, `.xlsm`, `.xlsb`, one with
      no Power Query at all, and one with a query whose M contains non-ASCII — that last one is the
      fixture that would have caught `FIX-3`.
- [x] ~~`src/refactor-settings.py` sat in `src/`.~~ Now `scripts/one-off/refactor-settings.py`,
      which is where a spent migration script belongs.

- [x] No `AGENTS.md` and no line-ending policy in the repo that every other repo is told to copy.
      Added, and 33 tracked files renormalized from CRLF.
- [x] ~~`release.yml` contained a corrupted byte (`EF BF BD`).~~ **Gone — verified zero occurrences
      2026-09-01.** The file was substantially rewritten for PQ-01 and PQ-34, which took it with it.
      [PQ-04]
- [x] ~~`.vsix` files in the repo root.~~ Four of them, in the end. Moved to `dist/`, and
      `package-vsix` now writes there — build output belongs with the other build output, not in the
      first place a human looks.
- [x] ~~`docs/archive/` holds 14 tracked files of v0.4.3 documentation.~~ **Deleted.** A hand-rolled
      archive is a second history surface in a repository that already has one; every file is
      recoverable from `106a2cd` and earlier. Archive folders are fine while iterating and should
      not be committed as a permanent record.
- [x] ~~`generate-expected-results.js` sat at the repo root.~~ Moved to `scripts/`.
- [x] ~~`docs/analysis/` and `docs/design/` do not exist here.~~ **Not true any more** — both exist
      and hold real work, exactly as intended: created when there was something to put in them.
- [ ] **Telemetry: decide, and write the decision down either way.** 5,500+ users and no telemetry,
      deliberately — which means breakage reaches us only when somebody opens an issue. That is not
      hypothetical: both bugs fixed this week reached real users, and neither was found by 142 tests
      or six CI legs. One surfaced from dogfooding at work, the other from a user who went back and
      reproduced his own problem after he had already solved it.

      What "add telemetry" actually involves, since the word does a lot of work: VS Code provides
      `env.isTelemetryEnabled` and a `TelemetryLogger`, and an extension is expected to honour the
      user's existing setting rather than ask separately. But the data has to land somewhere **we
      run** — Application Insights, typically — which means an endpoint, a key, a retention policy,
      and a privacy statement in both the README and the Marketplace listing. **The cost is the
      commitments, not the code.** Every one of those is a promise to 5,500 people, and promises are
      harder to withdraw than features.

      The narrower option, and currently the favourite: **no automatic reporting at all**, and a
      **Report an Issue** command that opens a prefilled GitHub issue carrying the version, the
      platform, and the last live-sync reason code. The user sees exactly what is being sent and
      chooses to send it. Most of the diagnostic value, none of the infrastructure, and no promise
      that outlives our interest in keeping it.

      It would also have collapsed a real round trip: a user hit a locked workbook, got a message
      blaming privilege levels, and it took a discussion thread to establish that the actual cause
      was a setting being off. One click carrying the reason code would have said so immediately.

---
