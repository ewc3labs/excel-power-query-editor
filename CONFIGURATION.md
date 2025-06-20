# Configuration Quick Reference

## Essential Settings

Access via `File` > `Preferences` > `Settings` > search "Excel Power Query"

### **Auto-Watch Setup** 
```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.verboseMode": true
}
```
*Automatically starts watching extracted files and shows detailed logs*

### **Custom Backup Location**
```json
{
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./PQ-backups",
  "excel-power-query-editor.maxBackups": 5
}
```
*Saves backups to custom folder, keeps 5 most recent*

### **Speed/Minimal Setup**
```json
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": false,
  "excel-power-query-editor.showStatusBarInfo": false
}
```
*Disables confirmations and backups for fast operation*

## Key Settings Explained

| Setting | What It Does | Recommended |
|---------|--------------|-------------|
| `watchAlways` | Auto-watch files when extracting | `true` for active development |
| `verboseMode` | Show detailed logs in Output panel | `true` for troubleshooting |
| `maxBackups` | Number of backup files to keep | `5-10` for development |
| `backupLocation` | Where to store backups | `"custom"` with organized path |
| `syncDeleteAlwaysConfirm` | Ask before deleting files | `true` for safety |

## Getting Verbose Output

1. Enable: `"excel-power-query-editor.verboseMode": true`
2. View: `View` > `Output` > select "Excel Power Query Editor"
3. See real-time logs of all operations

## Workspace vs User Settings

- **User Settings**: Apply everywhere
- **Workspace Settings**: Project-specific (`.vscode/settings.json`)

For Power Query projects, use workspace settings to auto-enable features for that project only.
