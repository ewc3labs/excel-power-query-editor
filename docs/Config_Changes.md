# Configuration Changes

Settings that were renamed or removed. **Worth reading when you upgrade.**

Where a change can be migrated automatically, it is — you should not have to do anything. This
document exists so that when a setting you recognize stops appearing where you expect it, there is
somewhere that says what happened to it.

Dates are approximate, newest first.

## Changes

**20260815 — every setting moved into a namespace.**

Thirteen settings were renamed. **Your existing values are migrated automatically the first time
0.6.0 starts**, in whichever scope you set them: user, workspace, or folder. A value you have
already set under the new name is never overwritten.

| Old | New |
| --- | --- |
| `autoBackupBeforeSync` | `backup.autoBackupBeforeSync` |
| `autoCleanupBackups` | `backup.autoCleanup` |
| `backupLocation` | `backup.location` |
| `customBackupPath` | `backup.customPath` |
| `logLevel` | `log.level` |
| `showStatusBarInfo` | `log.showStatusBarInfo` |
| `syncDeleteAlwaysConfirm` | `sync.deleteAlwaysConfirm` |
| `syncTimeout` | `sync.timeout` |
| `watchAlways` | `watch.always` |
| `watchAlwaysMaxFiles` | `watch.maxFiles` |
| `watchOffOnDelete` | `watch.offOnDelete` |
| `debugMode` | `log.level` set to `debug` |
| `verboseMode` | `log.level` set to `verbose` |

`debugMode` and `verboseMode` were booleans and became values of the `log.level` enum. If both were
set, `debug` wins. An explicit `log.level` you had already chosen wins over either.

The old names are still declared, marked deprecated, and will keep working for at least one more
release — so skipping a version does not skip the migration. They are struck through in the settings
UI with a pointer to the replacement.

> **If you ran a development build between 0.5.0 and 0.6.0**, an earlier migration deleted settings
> rather than migrating them. That code never shipped in a release, but if your configuration
> emptied itself, this is why. Setting them again is safe; the current migration preserves values.

**20260815 — Excel symbols are no longer written into your workspace.**

Previously the extension wrote `excel-pq-symbols.json` into `<workspace>/.vscode/` and appended that
folder to `powerquery.client.additionalSymbolsDirectories`. It now hands the symbols to the Power
Query extension directly through its API.

Nothing is written and no other extension's settings are modified. If you have an old
`.vscode/excel-pq-symbols/` folder, or an entry pointing at one in
`powerquery.client.additionalSymbolsDirectories`, **both are now unused and can be deleted**.
Neither does any harm if left.

See [Excel Symbols](Excel_Symbols.md).
