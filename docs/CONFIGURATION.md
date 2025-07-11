<!-- HEADER_TABLE -->
<table align="center">
<tr>
  <td width="112" align="center" valign="middle">
    <img src="assets/excel-power-query-editor-logo-128x128.png" width="128" height="128"><br>
    <strong>E · P · Q · E</strong>
  </td>

  <td align="center" valign="middle">
    <h1 align="center">Excel Power Query Editor</h1>
    <p align="left">
      <b>Edit Power Query M code directly from Excel files in VS Code. No Excel needed. No bullshit. It Just Works™.</b><br>
      <sub>
        Built by <strong>EWC3 Labs</strong> — where we rage-build the tools everyone needs, but nobody <del>cares to build</del>
        <em>is deranged enough to spend days perfecting until it actually works right.</em>
      </sub>
    </p>
  </td>

  <td width="112" align="center" valign="middle">
    <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128"><br>
    <strong><em>QA Officer</em></strong>
  </td>
</tr>
</table>
<!-- /HEADER_TABLE -->

---

## Configuration Reference

> **Complete settings guide with real-world use cases and optimization examples**

---

## 🚀 Quick Setup Commands

### Apply Recommended Defaults

**First time using the extension?** Run this command for optimal configuration:

```
Ctrl+Shift+P → "Excel Power Query: Apply Recommended Defaults"
```

**What it sets:**

- Auto-backup enabled with 5-file retention
- 500ms debounce delay (prevents CoPilot triple-sync)
- Watch mode ready but not auto-enabled
- Verbose logging disabled (clean experience)

## ⚙️ Complete Settings Reference

### Watch & Auto-Sync Settings

| Setting                     | Type    | Default | Description                             | Use Cases                                                                               |
| --------------------------- | ------- | ------- | --------------------------------------- | --------------------------------------------------------------------------------------- |
| `watchAlways`               | boolean | `false` | Auto-enable watch after extraction      | ✅ Active development<br>❌ Occasional editing                                          |
| `watchOffOnDelete`          | boolean | `true`  | Stop watching when `.m` file deleted    | ✅ Always recommended                                                                   |
| `sync.debounceMs`           | number  | `500`   | Delay before sync (prevents duplicates) | **300ms**: Fast workflows<br>**500ms**: CoPilot integration<br>**1000ms**: Slow systems |
| `watch.checkExcelWriteable` | boolean | `true`  | Verify Excel file access before sync    | ✅ Shared network drives<br>❌ Local SSD (performance)                                  |

**Example - Active Development:**

```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.sync.debounceMs": 300,
  "excel-power-query-editor.watch.checkExcelWriteable": true
}
```

**Example - CoPilot Integration:**

```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.sync.debounceMs": 500,
  "excel-power-query-editor.watch.checkExcelWriteable": true
}
```

### Sync Behavior Settings

| Setting                    | Type    | Default | Description                        | Use Cases                                                   |
| -------------------------- | ------- | ------- | ---------------------------------- | ----------------------------------------------------------- |
| `sync.openExcelAfterWrite` | boolean | `false` | Launch Excel after successful sync | ✅ Review changes immediately<br>❌ Automation/CI workflows |
| `syncDeleteAlwaysConfirm`  | boolean | `true`  | Confirm before "Sync and Delete"   | ✅ Safety (recommended)<br>❌ Trusted workflows only        |
| `syncTimeout`              | number  | `30000` | Sync operation timeout (ms)        | **15000ms**: Fast systems<br>**60000ms**: Large Excel files |

**Example - Interactive Workflow:**

```json
{
  "excel-power-query-editor.sync.openExcelAfterWrite": true,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": true,
  "excel-power-query-editor.syncTimeout": 30000
}
```

**Example - Automation/CI:**

```json
{
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": false,
  "excel-power-query-editor.syncTimeout": 60000
}
```

### Backup Management Settings

| Setting                | Type    | Default        | Description                           | Use Cases                                                                                  |
| ---------------------- | ------- | -------------- | ------------------------------------- | ------------------------------------------------------------------------------------------ |
| `autoBackupBeforeSync` | boolean | `true`         | Create backup before every sync       | ✅ Data protection<br>❌ SSD-constrained CI                                                |
| `backupLocation`       | enum    | `"sameFolder"` | Where to store backups                | **sameFolder**: Simple setup<br>**temp**: Clean workspace<br>**custom**: Organized storage |
| `customBackupPath`     | string  | `""`           | Custom backup directory               | `"./backups"`, `"../PQ-backups"`                                                           |
| `backup.maxFiles`      | number  | `5`            | Backup retention limit per Excel file | **3**: Minimal storage<br>**10**: Extensive history<br>**0**: Unlimited (not recommended)  |
| `autoCleanupBackups`   | boolean | `true`         | Auto-delete old backups               | ✅ Always recommended                                                                      |

**Example - Team Development:**

```json
{
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./project-backups",
  "excel-power-query-editor.backup.maxFiles": 10,
  "excel-power-query-editor.autoCleanupBackups": true
}
```

**Example - CI/CD Performance:**

```json
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.backupLocation": "temp",
  "excel-power-query-editor.backup.maxFiles": 2,
  "excel-power-query-editor.autoCleanupBackups": true
}
```

**Example - SSD-Constrained Environment:**

```json
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.backupLocation": "temp",
  "excel-power-query-editor.backup.maxFiles": 1
}
```

### Debug & Logging Settings

| Setting             | Type    | Default | Description                          | Use Cases                                                        |
| ------------------- | ------- | ------- | ------------------------------------ | ---------------------------------------------------------------- |
| `verboseMode`       | boolean | `false` | Detailed logs in Output panel        | ✅ Troubleshooting<br>✅ Understanding operations<br>❌ Clean UI |
| `debugMode`         | boolean | `false` | Debug-level logging + files          | ✅ Extension development<br>❌ Normal usage                      |
| `showStatusBarInfo` | boolean | `true`  | Show watch/sync status in status bar | ✅ Visual feedback<br>❌ Minimal UI                              |

**Example - Troubleshooting Setup:**

```json
{
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.debugMode": false,
  "excel-power-query-editor.showStatusBarInfo": true
}
```

**Example - Extension Development:**

```json
{
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.debugMode": true,
  "excel-power-query-editor.showStatusBarInfo": true
}
```

## 🎯 Configuration Scenarios

### Scenario 1: Solo Developer - Active Power Query Work

**Workflow:** Extract, edit, sync frequently with immediate Excel review

```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.sync.openExcelAfterWrite": true,
  "excel-power-query-editor.sync.debounceMs": 300,
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backup.maxFiles": 10,
  "excel-power-query-editor.verboseMode": false,
  "excel-power-query-editor.showStatusBarInfo": true
}
```

### Scenario 2: Team Development - Shared Project

**Workflow:** Multiple developers, version control, organized backups

```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.sync.debounceMs": 500,
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./team-backups",
  "excel-power-query-editor.backup.maxFiles": 15,
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": true
}
```

### Scenario 3: CI/CD Pipeline - Automated Processing

**Workflow:** Automated testing, performance-focused, minimal storage

```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.sync.debounceMs": 100,
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.backupLocation": "temp",
  "excel-power-query-editor.backup.maxFiles": 1,
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": false,
  "excel-power-query-editor.syncTimeout": 60000,
  "excel-power-query-editor.showStatusBarInfo": false
}
```

### Scenario 4: GitHub CoPilot Integration - Optimal AI Workflow

**Workflow:** CoPilot-assisted development with intelligent sync prevention

```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.sync.debounceMs": 500,
  "excel-power-query-editor.watch.checkExcelWriteable": true,
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backup.maxFiles": 8,
  "excel-power-query-editor.verboseMode": false,
  "excel-power-query-editor.showStatusBarInfo": true
}
```

### Scenario 5: Large Excel Files - Performance Optimized

**Workflow:** Working with multi-MB Excel files, prioritizing speed

```json
{
  "excel-power-query-editor.sync.debounceMs": 1000,
  "excel-power-query-editor.watch.checkExcelWriteable": false,
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backupLocation": "temp",
  "excel-power-query-editor.backup.maxFiles": 3,
  "excel-power-query-editor.syncTimeout": 120000,
  "excel-power-query-editor.verboseMode": true
}
```

## 🔧 Settings Organization

### User Settings vs Workspace Settings

**User Settings** (`File > Preferences > Settings`):

- Applied globally across all VS Code projects
- Good for personal preferences (UI, logging, default behavior)

**Workspace Settings** (`.vscode/settings.json`):

- Applied only to current project
- Perfect for team collaboration and project-specific configurations
- Committed to version control for team consistency

### Recommended Split:

**User Settings** (Personal Preferences):

```json
{
  "excel-power-query-editor.verboseMode": false,
  "excel-power-query-editor.showStatusBarInfo": true,
  "excel-power-query-editor.sync.openExcelAfterWrite": true
}
```

**Workspace Settings** (Project Configuration):

```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.sync.debounceMs": 500,
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./project-backups",
  "excel-power-query-editor.backup.maxFiles": 10
}
```

## 📋 Migration Guide: v0.4.x → v0.5.0

### New Settings in v0.5.0:

- `sync.openExcelAfterWrite` - Automatically open Excel after sync
- `sync.debounceMs` - Configurable sync delay (prevents CoPilot triple-sync)
- `watch.checkExcelWriteable` - Excel file access validation
- `backup.maxFiles` - Replaces deprecated `maxBackups`

### Deprecated Settings:

- `syncDeleteTurnsWatchOff` - Functionality merged with `watchOffOnDelete`

### Automatic Migration:

The extension automatically migrates your v0.4.x settings. **No action required.**

### Manual Migration (Optional):

```json
// v0.4.x
{
  "excel-power-query-editor.maxBackups": 5
}

// v0.5.0 (improved)
{
  "excel-power-query-editor.backup.maxFiles": 5
}
```

## 🔍 Accessing Settings

### Via VS Code UI:

1. `File > Preferences > Settings` (or `Ctrl+,`)
2. Search for `"Excel Power Query"`
3. Configure settings with UI controls

### Via settings.json:

1. `Ctrl+Shift+P` → "Preferences: Open Settings (JSON)"
2. Add your configuration
3. IntelliSense provides auto-completion

### Via Command Palette:

```
Ctrl+Shift+P → "Excel Power Query: Apply Recommended Defaults"
```

## 🚨 Troubleshooting Configuration

### Settings Not Taking Effect:

1. **Reload VS Code**: `Ctrl+Shift+P` → "Developer: Reload Window"
2. **Check settings scope**: User vs Workspace settings priority
3. **Validate JSON syntax**: Ensure proper formatting in settings.json

### Performance Issues:

1. **Reduce backup retention**: Lower `backup.maxFiles`
2. **Increase debounce delay**: Higher `sync.debounceMs`
3. **Disable unnecessary features**: Turn off `sync.openExcelAfterWrite`

### Debug Configuration Issues:

```json
{
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.debugMode": true
}
```

Then check: `View > Output > "Excel Power Query Editor"`

## 🔗 Related Documentation

- **📖 [User Guide](USER_GUIDE.md)** - Complete workflows and feature explanations
- **🤝 [Contributing](CONTRIBUTING.md)** - Development setup and testing configuration
- **📝 [Changelog](../CHANGELOG.md)** - Version history and setting changes

---

**Excel Power Query Editor v0.5.0** - Professional configuration for every workflow
