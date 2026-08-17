# Overview

Documentation for the Excel Power Query Editor extension.

The structure follows [Klipper][klipper]: flat files, one document per feature, and two references
that list everything. A feature arrives with its own document and an entry in the references, so
adding one does not mean editing a doc somebody else owns.

## Getting started

- [Installation](Installation.md): installing the extension and what it needs.
- [User Guide](User_Guide.md): extracting, editing and syncing — the whole workflow.
- [FAQ](FAQ.md): the questions that come up most.
- [Beta Downloads](Beta_Downloads.md): pre-release builds, ahead of the Marketplace.

## Configuration and reference

- [Config Reference](Config_Reference.md): every setting, generated from the source so it cannot
  drift.
- [Commands](Commands.md): every command in the palette and the context menus.
- [Config Changes](Config_Changes.md): settings that were renamed or removed. **Read this when
  upgrading.**

## Features

- [Live Sync](Live_Sync.md): writing to a workbook Excel already has open.
- [Watch Mode](Watch_Mode.md): syncing automatically when you save.
- [Backups](Backups.md): what is kept before a workbook is written, and for how long.
- [Excel Symbols](Excel_Symbols.md): IntelliSense for `Excel.CurrentWorkbook()` and friends.

## Developer documentation

- [Contributing](CONTRIBUTING.md): setting up, testing, and sending a change.
- [Publishing](PUBLISHING_GUIDE.md): how a release is cut.
- [Design notes](design/): why things work the way they do, and what was rejected.
- [Analysis](analysis/): measurements, and the findings behind particular decisions.
- [Project](project/): the roadmap, the slice registry, and the check-in punchlist.
- [RAG sessions](RAG_Sessions/): worked sessions kept for their reasoning, not their conclusions.

## Adding documentation

If you are adding a feature:

1. Write `Your_Feature.md` — what it does, how to turn it on, and how it fails.
2. Add a line to the **Features** list above.
3. Settings appear in [Config Reference](Config_Reference.md) automatically; run
   `npm run docs:config` after changing `package.json`.
4. Commands go in [Commands](Commands.md) by hand.
5. If you renamed or removed a setting, add a dated entry to [Config Changes](Config_Changes.md).

Titles are `Title_Case_With_Underscores.md`. Keep documents about one thing; a document that needs
two titles is two documents.

**Except the SHOUTING ones.** `README.md`, `CHANGELOG.md`, `LICENSE`, `SUPPORT.md`,
`CONTRIBUTING.md` and `PUBLISHING_GUIDE.md` keep their all-caps names. Those are "how we work here"
documents rather than documentation of a feature - the convention is older than this repository,
GitHub gives some of them special treatment, and people look for them by that name. Feature docs get
title case; institutional docs shout.

`npm run docs:links` checks two things, and CI runs it alongside `npm run docs:check`:

- **Every link resolves**, with the case it was written in. Case matters because Windows and macOS
  resolve a wrong-case link happily and GitHub does not, so it works on every machine that could
  catch it and breaks for every reader.
- **Every document is reachable.** A file nothing links to is invisible, and stays correct and
  unread until it quietly goes stale. Link it here, or delete it.

[klipper]: https://github.com/Klipper3d/klipper/tree/master/docs
