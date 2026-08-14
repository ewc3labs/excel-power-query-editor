# modules

One document per subsystem, for things too big to be a slice.

`src/extension.ts` is currently 2,152 lines and contains every subsystem this extension has. As pieces
are extracted — workbook read/write, watching, backups, configuration, commands — each earns a document
here describing what it owns, what it must never do, and how it fails.

The read/write path is the one that matters most: it is the only code in this repo that can destroy
somebody's work.

Name the file after the subsystem: `Workbook_ReadWrite.md`.
