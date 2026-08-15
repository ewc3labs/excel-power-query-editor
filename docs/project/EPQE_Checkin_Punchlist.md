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

---

## 2026-08-14 — first pass, adopting the HQ conventions

**Items:**

- [ ] **The dev loop produces false negatives.** Reinstalling the extension does not affect a
      RUNNING VS Code: the extension host keeps executing whatever it loaded at startup, so a fix can
      be built, packaged, installed and verified on disk while the editor still runs the old code.
      That cost a full debugging cycle on live sync - the code was correct and the symptom said
      otherwise. Worse, `package.json` stays at the same version through dozens of dev installs, and
      VS Code keys the extension folder by version, so same-version reinstalls are not reliably a
      clean replace either. `npm run bump-version` exists and is not being used in the loop.
      Wanted: a dev-install script that bumps a prerelease number, packages, installs, and either
      reloads the window or says plainly that a reload is required.
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
      (`extension.ts:306`). It enumerates the config and sets every key to `undefined` in both User and
      Workspace scope, preserving nothing, and its guard compares against the *extension version* so it
      re-fires on every bump. The correct earlier attempt sits commented out directly above it. Full
      write-up in `docs/analysis/settings-migration.md`. [PQ-09]
- [x] **Why removing old settings was impossible:** VS Code only lets an extension write a key that is
      *registered*. Once the old names were deleted from `contributes.configuration`, they could no
      longer be cleared — the value sits in the user's settings.json permanently as an unknown setting.
      The fix is to keep them declared with `deprecationMessage`, migrate, then clear. [PQ-09]
- [ ] The code still references the OLD setting names — `logLevel` 13 times, `debugMode` 9,
      `verboseMode` 9, `backupLocation` 3, `watchAlways` 2 — while `package.json` declares only the
      new ones. Those reads can only ever return defaults. Needs a proper audit, not a sweep. [PQ-09]
- [x] **The README swap:** `docs/README.gh.md` and `docs/README.vsmarketplace.md` were both **0 bytes**,
      emptied in `a2ea1ef` ("save before settings refactor") — a work-in-progress commit. Recovered
      from `fb52fc0` (v0.5.0): 8,034 and 4,034 bytes respectively.
- [ ] `scripts/set-readme-gh.js` and `set-readme-vsce.js` still exist and **nothing calls them** — not
      package.json, not the workflows. Left as-is they are a loaded weapon: running either copies its
      source over `README.md`, and until just now both sources were empty. Decide one README or two,
      then delete whichever machinery loses. [PQ-10]
- [ ] EPQE writes `excel-pq-symbols.json` into `.vscode/` in the workspace root, which is somebody
      else's folder. Worth researching whether Microsoft's PQ/M extension now ships these symbols, and
      whether there is a way to reference them that does not require writing into a user's workspace at
      all. Also: what happens when a plain folder is opened with no workspace. [PQ-11]
- [ ] `test/fixtures/test_workbook.xlsx` and `test_workbook.xlsb` are not workbooks — the first
      contains the literal text "test xlsx file". `test_clean.xlsx` is four bytes. Byte-identical to
      v0.5.0, so they have always been stubs; worth knowing before trusting a test that opens one.
- [ ] `src/refactor-settings.py` was committed into `src/` during the refactor. A one-off migration
      script does not live beside the extension source.

- [x] No `AGENTS.md` and no line-ending policy in the repo that every other repo is told to copy.
      Added, and 33 tracked files renormalised from CRLF.
- [ ] `release.yml:260` contains a corrupted byte — `EF BF BD` (replacement character) in a step name.
      Cosmetic on its own; a fair signal about how carefully the file has been reviewed. [PQ-04]
- [ ] Three `.vsix` files sit in the repo root. Untracked and gitignored, so harmless, but they are
      build output living where a human looks first.
- [ ] `docs/archive/` holds 14 tracked files of v0.4.3 documentation. Worth deciding whether that is
      history worth carrying or clutter to delete.
- [ ] `generate-expected-results.js` sits at the repo root while everything else of its kind is in
      `scripts/`.
- [ ] `docs/analysis/` and `docs/design/` do not exist here, though the org conventions expect them.
      Not worth creating empty — create when there is something to put in them.
- [ ] Org-wide: `RAG_sessions` (lowercase) in 8 repos, `RAG_Sessions` in 2. HQ declares the capital
      canonical and is outvoted. Cheaper to move the standard than eight folders — needs a decision.
- [ ] The extension has real users and no telemetry, deliberately. That means we learn about breakage
      only from issues. Worth deciding whether that stays true forever, in writing, either way.
