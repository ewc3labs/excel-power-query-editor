# Settings migration — why it fought back, and what actually works

**Status:** fixed and verified 2026-08-14 (`PQ-09`). This document is kept because the trap that
caused it is not obvious and will be met again by anyone renaming a setting.

## What is in the repo right now

`src/extension.ts:306` calls `migrateLegacySettings()` during activation. That function does this:

```ts
async function wipeConfig(config, scope) {
    const keys = Object.keys(config);
    for (const key of keys) {
        await config.update(key, undefined, scope);   // delete
    }
    await config.update(migrationKey, extensionVersion, scope);
}
```

It reads no values and preserves nothing. It enumerates the configuration object and sets every key
to `undefined`, in both User and Workspace scope. The guard is `migrationKey !== extensionVersion`,
so it re-runs **on every version bump**, not once.

The extension has thousands of installs. Publishing this deletes their configuration on first
launch.

There is also a commented-out earlier attempt directly above it (`extension.ts:1707-1746`) which is
the *right shape* — read `debugMode`, decide a `log.level`, write the new value, clear the old ones
— and it was abandoned. The wipe replaced it.

## Why the honest version was hard

0.5.2 renamed 13 of 19 settings into namespaces and removed `debugMode` and `verboseMode` outright,
then deleted the old names from `contributes.configuration` in `package.json`.

**That is the trap.** VS Code will only let an extension write a configuration key that is
*registered*. Once a key is gone from `contributes.configuration`,
`config.update(oldKey, undefined, scope)` fails — the key is no longer known, so it cannot be
cleared. The old value stays in the user's `settings.json` forever, grayed out as an unknown
setting, and there is no API to remove it.

So the sequence "rename the settings, then migrate" cannot work in that order. The rename removes
the very handle the migration needs.

## What works

**Keep the old keys declared, marked deprecated.** They stay registered, so they can still be read
and cleared, and VS Code renders them struck through with an explanation:

```json
"excel-power-query-editor.watchAlways": {
  "type": "boolean",
  "default": false,
  "markdownDeprecationMessage": "Use `#excel-power-query-editor.watch.always#` instead.",
  "deprecationMessage": "Deprecated: use excel-power-query-editor.watch.always"
}
```

Then migrate explicitly, per key, per scope:

1. `inspect()` the old key to find out **which scope** actually holds a value — `globalValue`,
   `workspaceValue`, `workspaceFolderValue`. Do not guess, and do not write to a scope the user
   never used, or the setting appears in a file they did not choose.
2. If the old key has a value **and the new key does not**, write the new one in that same scope.
   Never overwrite a new value the user has already set.
3. Clear the old key in that scope only.
4. Record that migration ran, keyed to a **migration schema number** — not the extension version.
   Version-keyed markers re-run on every release, which is how the current code became a repeated
   wipe rather than a one-off.

Booleans that collapsed into an enum (`debugMode`/`verboseMode` → `log.level`) need a precedence
rule written down: debug beats verbose, and neither beats an explicit `log.level` the user has set.

## Leave the deprecated keys in for at least one minor release

They cost nothing but a few lines of JSON, they let a user who skipped a version still migrate, and
removing them early recreates exactly this problem for whoever is here next.

## Verified, on a real install

The wipe was not hypothetical: the marker `xtn.level: "0.5.2"` and a stale `settings.level: "0.5.2"`
were sitting in the developer's own VS Code settings, with everything else for this extension gone.
It had already run.

The replacement was tested end to end rather than only in unit tests — seven v0.5.0-era settings
were written into a real `settings.json`, the built `.vsix` installed with `--force`, and VS Code
started:

| seeded | became |
| --- | --- |
| `watchAlways: true` | `watch.always: true` |
| `watchAlwaysMaxFiles: 42` | `watch.maxFiles: 42` |
| `syncTimeout: 15` | `sync.timeout: 15` |
| `backupLocation: "customFolder"` | `backup.location: "customFolder"` |
| `customBackupPath: "C:/Temp/pq-backups"` | `backup.customPath: "C:/Temp/pq-backups"` |
| `debugMode: true` | `log.level: "debug"` |
| `autoCleanupBackups: true` | `backup.autoCleanup: true` |

All seven old keys were cleared. `watch.checkExcelWriteable`, which the migration has no business
touching, was still there afterwards — the specific thing the wipe destroyed. A second activation
changed the file not at all: the schema marker holds, so this cannot become a repeating wipe the way
a version-keyed marker did.
