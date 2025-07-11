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

## Complete User Guide

> **Power Query productivity unlocked for engineers and data nerds**

---

## 🚀 The Complete Workflow: Extract → Edit → Watch → Sync

### 1. 📤 Extract Power Query from Excel

**Right-click any Excel file** (`.xlsx`, `.xlsm`, `.xlsb`) in VS Code Explorer:

```
my-data-analysis.xlsx  ← Right-click here
└── Select "Extract Power Query from Excel"
```

**What happens:**

- Extension reads Power Query queries from Excel's internal DataMashup
- Creates `my-data-analysis.xlsx_PowerQuery.m` file
- Opens automatically with full M language syntax highlighting
- **Preserves all comments and formatting**

**Supported Excel formats:**

- `.xlsx` - Standard Excel workbook
- `.xlsm` - Macro-enabled workbook
- `.xlsb` - Binary workbook (faster for large files)

### 2. ✏️ Edit with Full VS Code Power

**IntelliSense & Syntax Highlighting:**

- Install recommended: `powerquery.vscode-powerquery` (auto-installed)
- Full M language support with autocomplete
- Error highlighting and syntax validation
- Comment preservation through sync cycles

**Pro Tips:**

- Use `Ctrl+/` for quick commenting
- `F12` for function definitions (with Power Query extension)
- `Ctrl+Shift+P` → "Format Document" for clean code

### 3. 👁️ Enable Auto-Watch (Recommended)

**Right-click your `.m` file** → **"Toggle Watch"** or **"Watch File for Changes"**

**Status Bar Indicator:**

```
👁 Watching 1 PQ file    ← Shows active watch count
```

**What Auto-Watch Does:**

- Monitors `.m` file for saves (`Ctrl+S`)
- **Intelligent debouncing** prevents duplicate syncs (configurable 500ms delay)
- Automatic backup before each sync
- **Smart change detection** - only syncs when content actually changes

### 4. 🔄 Sync Changes Back to Excel

**Automatic (with Watch enabled):**

- Save your `.m` file (`Ctrl+S`)
- Watch triggers sync automatically
- Backup created, Excel updated
- **Optional**: Automatically opens Excel after sync

**Manual Sync:**

- Right-click `.m` file → **"Sync Power Query to Excel"**
- Useful for one-off changes without enabling watch

**Sync Process:**

1. **Backup Creation**: `my-data-analysis_backup_2025-07-11_14-30-45.xlsx`
2. **Content Validation**: Ensures M code is syntactically valid
3. **Excel Update**: Replaces Power Query content in Excel's DataMashup
4. **Verification**: Confirms successful write operation

## 🛠️ Advanced Features & Configuration

### Smart Defaults (v0.5.0)

**First-time setup?** Run this command:

- `Ctrl+Shift+P` → **"Excel Power Query: Apply Recommended Defaults"**

**Sets optimal configuration for:**

- Auto-backup enabled
- 500ms debounce delay
- Watch mode behavior
- Backup retention (5 files)

### Backup Management

**Automatic Backups:**

- Created before every sync operation
- Timestamped: `filename_backup_YYYY-MM-DD_HH-MM-SS.xlsx`
- Configurable retention limit (default: 5 files per Excel file)
- Auto-cleanup when limit exceeded

**Backup Locations:**

- **Same Folder** (default): Next to original Excel file
- **System Temp**: OS temporary directory
- **Custom Path**: Specify your own backup directory

**Manual Cleanup:**

- `Ctrl+Shift+P` → **"Excel Power Query: Cleanup Old Backups"**
- Select Excel file to clean up its backups

### CoPilot Integration (v0.5.0 Excellence)

**Problem Solved:** CoPilot Agent mode causing triple-sync

- ✅ **Intelligent Debouncing**: 500ms delay prevents duplicate operations
- ✅ **File Hash Deduplication**: Only syncs when content actually changes
- ✅ **Smart Change Detection**: Timestamp + content validation

**Optimal CoPilot Workflow:**

1. Enable watch mode on your `.m` file
2. Use CoPilot to edit/refactor your Power Query
3. Accept CoPilot suggestions
4. Single sync triggered automatically (no duplicates!)

### Team Collaboration Best Practices

**Shared Projects:**

```json
// .vscode/settings.json (workspace)
{
  "excel-power-query-editor.autoBackupBeforeSync": true,
  "excel-power-query-editor.backup.maxFiles": 10,
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.verboseMode": true
}
```

**CI/CD Integration:**

```json
// Disable interactive features for automation
{
  "excel-power-query-editor.sync.openExcelAfterWrite": false,
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.watchAlways": false
}
```

**Performance Optimization:**

```json
// For SSD-constrained environments
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.backupLocation": "temp"
}
```

## 🔧 All Available Commands

Access via `Ctrl+Shift+P` (Command Palette) or right-click context menus:

| Command                            | Context                | Description                      |
| ---------------------------------- | ---------------------- | -------------------------------- |
| **Extract Power Query from Excel** | Right-click Excel file | Extract queries to `.m` file     |
| **Sync Power Query to Excel**      | Right-click `.m` file  | Manual sync back to Excel        |
| **Watch File for Changes**         | Right-click `.m` file  | Enable auto-sync on save         |
| **Toggle Watch**                   | Right-click `.m` file  | Toggle watch on/off              |
| **Stop Watching**                  | Right-click `.m` file  | Disable auto-sync                |
| **Sync and Delete**                | Right-click `.m` file  | Sync then delete `.m` file       |
| **Raw Extract (Debug)**            | Right-click Excel file | Debug extraction with raw output |
| **Apply Recommended Defaults**     | Command Palette        | Set optimal configuration        |
| **Cleanup Old Backups**            | Command Palette        | Manual backup management         |

## 🚨 Troubleshooting

### Excel File Locked / Permission Issues

**Problem:** "Could not sync - Excel file is locked"

**Solutions:**

1. **Close Excel** if file is open in Excel application
2. **Check file permissions** - ensure VS Code can write to the file
3. **Wait and retry** - Excel may release lock after a moment
4. **Enable write access checking**: Set `watch.checkExcelWriteable: true` (default)

**Advanced:** Extension automatically detects locked files and retries with user feedback.

### Right-Click Menu Not Working

**Problem:** Context menu commands not appearing

**Solutions:**

1. **Click inside the editor** - Commands require editor focus, not just file selection
2. **Reload VS Code** - `Ctrl+Shift+P` → "Developer: Reload Window"
3. **Check file type** - Ensure `.xlsx/.xlsm/.xlsb` for Excel files, `.m` for Power Query files
4. **Extension activation** - Commands appear after first usage

### Sync Failures / Corrupted Excel Files

**Problem:** Sync operation fails or produces corrupted Excel

**Solutions:**

1. **Restore from backup** - Use timestamped backup files
2. **Validate M syntax** - Ensure your Power Query code is syntactically correct
3. **Enable verbose logging**: Set `verboseMode: true` in settings
4. **Check Output panel**: View → Output → "Excel Power Query Editor"

**Debug Mode:**

```json
{
  "excel-power-query-editor.debugMode": true,
  "excel-power-query-editor.verboseMode": true
}
```

### Watch Mode Not Triggering

**Problem:** Auto-sync not working when saving `.m` file

**Diagnostic Steps:**

1. **Check status bar** - Look for `👁 Watching X PQ files`
2. **File saved?** - Ensure you actually saved (`Ctrl+S`) the `.m` file
3. **Debounce delay** - Wait 500ms after save (configurable)
4. **Toggle watch** - Right-click → "Toggle Watch" to restart

**Common Causes:**

- File system events not firing (rare)
- Debounce period too long for your workflow
- Excel file locked preventing sync

### Performance Issues

**Problem:** Slow sync operations or VS Code lag

**Solutions:**

1. **Reduce backup retention**: Lower `backup.maxFiles` setting
2. **Use temp folder for backups**: Set `backupLocation: "temp"`
3. **Disable auto-open Excel**: Set `sync.openExcelAfterWrite: false`
4. **Increase debounce delay**: Higher `sync.debounceMs` for rapid saves

**Large Excel Files:**

- Use `.xlsb` format for better performance
- Consider disabling auto-backup for CI/CD scenarios
- Monitor backup disk usage with high retention settings

## ⌨️ Keyboard Shortcuts & Power User Tips

### Efficient Workflows

**Extract Multiple Files:**

```bash
# Use VS Code's multi-select (Ctrl+click) then right-click → Extract
File1.xlsx  ← Ctrl+click
File2.xlsx  ← Ctrl+click
File3.xlsx  ← Right-click → "Extract Power Query from Excel"
```

**Bulk Watch Setup:**

```bash
# After extraction, select all .m files → Right-click → "Watch File for Changes"
*.m files → Ctrl+A → Right-click → Enable watch
```

**Quick Configuration:**

```bash
# Command Palette shortcuts
Ctrl+Shift+P → "excel" → Shows all extension commands
Ctrl+Shift+P → "apply" → Quick access to Apply Recommended Defaults
```

### Status Bar Integration

Monitor your Power Query workflow:

```
👁 Watching 3 PQ files    ← Active watch count
🔄 Syncing...             ← Sync in progress
✅ Synced to Excel        ← Successful sync (temporary)
❌ Sync failed           ← Error occurred (temporary)
```

**Click status bar** for quick actions and detailed information.

### Integration with Other Extensions

**Recommended Extension Stack:**

```vscode-extensions
powerquery.vscode-powerquery        # M language support (auto-installed)
ms-vscode.vscode-json               # Settings.json editing
eamodio.gitlens                     # Git integration for .m files
ms-python.python                    # For data analysis workflows
```

**Git Integration:**

- Add `.m` files to version control
- `.gitignore` backup files: `*_backup_*.xlsx`
- Use Git diff to review Power Query changes

## 🔗 Related Documentation

- **⚙️ [Configuration Reference](CONFIGURATION.md)** - Complete settings guide with examples
- **🤝 [Contributing Guide](CONTRIBUTING.md)** - Development setup and contribution guidelines
- **📝 [Changelog](../CHANGELOG.md)** - Version history and feature updates

---

**Excel Power Query Editor v0.5.0** - Professional-grade Power Query development in VS Code
