# AGENTS.md — Excel Power Query Editor

Project instructions. The org baseline is `ewc3labs-hq/AGENTS.md` — identity, method, tone, and the
conventions every repo shares. This file covers what is different **here**, and this repo is different
in one way that outranks everything else: it writes to other people's spreadsheets.

---

## What this actually is

A VS Code extension that edits Power Query **M** code inside `.xlsx` / `.xlsm` / `.xlsb` files with no
Excel, no COM, and no Windows requirement. It unzips the workbook, finds the DataMashup part, decodes
it, hands you the M as a normal editable file, and writes it back.

**It has real users.** This is the most-installed thing EWC3 Labs has shipped, it is on the VS Code
Marketplace, and people run it against workbooks they did not back up. Every decision in this repo is
downstream of that.

It is also the reference implementation for org conventions — `src/` + `test/` + `docs/`, esbuild,
eslint flat config, `.github/workflows/{ci,release}.yml`, issue templates. Other repos are told to
copy this one, so sloppiness here propagates.

---

## The stakes: never corrupt a workbook

The unrecoverable failure is destroying someone's file. Not crashing — **corrupting**, quietly, in a
workbook holding work nobody can reconstruct.

- **Back up before every write.** That is what the backup settings exist for. Do not add a path that
  writes to a workbook without one, however small the change looks.
- **A failed write must leave the original intact.** Write to a temp file and swap; never edit in
  place and hope.
- **`.xlsb` is binary and unforgiving.** Treat any change touching it as high-risk and test on real
  files, not fixtures alone.
- If a change alters how bytes reach a workbook, say so plainly in the changelog. Users deciding
  whether to upgrade are deciding whether to risk their data.

The library doing the hard part is [`excel-datamashup`](https://github.com/Vladinator/excel-datamashup).
We stand on it, and on [EditExcelPQM](https://github.com/amalanov/EditExcelPQM). Both stay credited.

---

## Architecture — be honest about the monolith

```
src/extension.ts      2,152 lines   ← everything
src/configHelper.ts      80 lines
```

That is the whole extension. Ten commands, eighteen settings, file watching, backup handling, symbol
installation and the Excel read/write path all live in one file.

**This is a known problem, not a style someone chose.** When you touch it:

- Do not rewrite it wholesale because it offends you. It works, it has users, and a big-bang refactor
  of an undertested monolith is how you corrupt somebody's workbook.
- Do extract the piece you are already working in, with tests, and leave it better than you found it.
  The natural seams are: workbook read/write, watch, backup, config, commands.
- New functionality of any size gets its own module. Do not add to the pile.

`test/` is real — roughly 2,900 lines across backup, commands, integration, watch and utils, running
under `vscode-test`. Keep it that way; it is the only thing standing between a refactor and a
corrupted file.

---

## Build, test, ship

```bash
npm run compile        # check-types + lint + esbuild
npm test               # vscode-test — launches a real VS Code
npm run package        # production bundle
npm run package-vsix   # .vsix for local install
npm run dev-install     # build + install into your VS Code, forced
```

`npm test` needs a display and downloads VS Code on first run. It is slower than it looks; it is not
hung.

---

## Release pipeline — currently the weakest part

The intent is right: CI validates, release builds the `.vsix`, publishes a GitHub Release, and pushes
stable tags to the Marketplace with `VSCE_PAT`. In practice it is flaky, and there are concrete
reasons rather than mysteries:

- **`release.yml` is triggered by `workflow_run` off CI**, not by the tag. That pattern is fragile: it
  runs the workflow file from the default branch rather than the tag, the triggering context is
  awkward to recover, and failures are hard to reproduce. **Tag-triggered is the fix** — see
  `ewc3-recall-tape/.github/workflows/release.yml`, which is a third the size and has not missed.
- **318 lines** for a release is a lot of surface. Most of it is reporting.
- There is a **corrupted byte in a step name** around `release.yml:260` (`�🚀`), which is a decent
  signal for how carefully the rest has been reviewed.

**State of play worth knowing before you touch versioning:** `package.json` says **0.5.2**, and the
newest published release is **v0.5.0 (2025-07-21)**. There are `.vsix` files for 0.5.2 sitting in the
repo root. Work has been built and never shipped. Do not assume the version in `package.json` is live.

Also: `.vsix` build artifacts do not belong in the repo root.

---

## Repo standards

- **The Marketplace listing is the product page.** README, screenshots/GIFs, categories and keywords
  are features. `docs/README.vsmarketplace.md` and `docs/README.gh.md` exist because the two audiences
  differ — keep them in step.
- **SEO is a feature.** People search `edit power query vscode`, `power query m editor`, `excel
  without excel`. Name for humans, metadata for engines.
- **Settings are a contract.** Eighteen of them are already public. Renaming or removing one breaks
  configs silently, so deprecate and migrate rather than rename.
- **Credit generously**, and keep the acknowledgements section current.
- **No telemetry** without an explicit, documented, opt-in decision from Wilson.

---

## Conventions inherited from the org

- `docs/analysis/` evidence · `docs/design/` architecture and *why* · `docs/RAG_sessions/` how a hard
  problem got solved
- Commit messages carry the **why**; the diff already has the what
- **Don't commit without Wilson reviewing the diff**, and don't push unless asked
- LF everywhere — `.gitattributes` and `.editorconfig` in every repo; check with `git ls-files --eol`
- Skills load from `ewc3labs-hq/.claude/skills/` in every repo: roadmap, RAG sessions, documentation
  standards, new machine

---

## When in doubt

Ask what happens to a user's workbook if this goes wrong halfway through. If the answer is anything
other than "nothing," that is the thing to fix first.
