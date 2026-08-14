# Excel Power Query Editor - User Guide

## Overview
This VS Code extension provides a modern, reliable way to extract Power Query M code from Excel files, edit it with full VS Code functionality, and sync changes back to Excel. No COM dependencies, no Excel installation required, and works across platforms.

## 🚀 Quick Start

### 1. Install the Extension(s)
**This extension requires the Microsoft Power Query / M Language extension:**

```vscode-extensions
powerquery.vscode-powerquery
```

**Install from VS Code Marketplace (Recommended):**

1. **Extensions View**: Open VS Code → Extensions (`Ctrl+Shift+X`) → Search "Excel Power Query Editor" → Install
2. **Command Line**: `code --install-extension ewc3labs.excel-power-query-editor`
3. **Direct Link**: [VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor)

**Alternative - VSIX File**: `code --install-extension excel-power-query-editor-[version].vsix`

*The Power Query extension will be automatically installed via Extension Pack.*

### 2. Extract Power Query from Excel
1. Right-click any `.xlsx`, `.xlsm`, or `.xlsb` file in Explorer
2. Select **"Extract Power Query from Excel"**
3. Extension creates `filename.xlsx_PowerQuery.m` in the same directory
4. File opens automatically with syntax highlighting

### 3. Edit Your Power Query Code
- Full VS Code editing experience with IntelliSense
- Syntax highlighting for M language
- Comments preserved during sync operations
- Save changes normally (`Ctrl+S`)

### 4. Sync Changes Back to Excel
**Manual Sync:**
1. Right-click the `.m` file
2. Select **"Sync Power Query to Excel"**
3. Automatic backup created, changes applied

**Auto-Sync (Recommended):**
1. Right-click the `.m` file
2. Select **"Watch File for Changes"** or **"Toggle Watch"**
3. Any saved changes automatically sync to Excel
4. Status bar shows `👁 Watching X PQ files`

## 📋 All Available Commands

### Context Menu Commands (Right-Click)

#### **On Excel Files** (`.xlsx`, `.xlsm`, `.xlsb`):
- **Extract Power Query from Excel** - Create `.m` files from Power Query
- **Raw Excel Extraction (Debug)** - Extract all Excel components for debugging
- **Cleanup Old Backups** - Manage backup files for this Excel file

#### **On Power Query Files** (`.m`):
- **Sync Power Query to Excel** - Update Excel with current `.m` file content
- **Watch File for Changes** - Enable automatic sync on file save
- **Toggle Watch** - Smart toggle: start watching if not watched, stop if watched
- **Sync & Delete** - Sync to Excel then safely delete the `.m` file

### Command Palette (`Ctrl+Shift+P`)
All commands available via Command Palette with `Excel PQ:` prefix:
- `Excel PQ: Extract Power Query from Excel`
- `Excel PQ: Sync Power Query to Excel` 
- `Excel PQ: Watch File for Changes`
- `Excel PQ: Toggle Watch`
- `Excel PQ: Stop Watching File`
- `Excel PQ: Sync & Delete`
- `Excel PQ: Raw Excel Extraction (Debug)`
- `Excel PQ: Cleanup Old Backups`

## 📁 File Naming Convention

The extension uses a **full filename** approach for better organization:

### **Naming Pattern**:
- **Excel file**: `MyWorkbook.xlsx`
- **Power Query file**: `MyWorkbook.xlsx_PowerQuery.m`

### **Examples**:
```
Financial_Report.xlsx          → Financial_Report.xlsx_PowerQuery.m
SalesData.xlsm                → SalesData.xlsm_PowerQuery.m
Dashboard.xlsb                 → Dashboard.xlsb_PowerQuery.m
Q4_Analysis_2025.xlsx          → Q4_Analysis_2025.xlsx_PowerQuery.m
```

### **Auto-Detection Logic**:
The sync feature finds Excel files by:
1. **Removing** `_PowerQuery.m` from filename
2. **Checking** for exact match: `filename.xlsx`, `filename.xlsm`, `filename.xlsb`
3. **Searching** same directory first, then parent directories
4. **Prompting** for manual selection if not found

## 🔄 Auto-Watch Feature

### **What It Does**:
- Monitors `.m` files for changes
- Automatically syncs to Excel on save
- Survives VS Code reloads (if **Watch Always** setting enabled)
- Shows status in status bar

### **How to Use**:
1. **Enable Watch Always**: `Settings` → `Excel-power-query-editor: Watch Always`
2. **Extract any file** → Automatically starts watching
3. **Or manually**: Right-click `.m` file → "Toggle Watch"

### **Status Indicators**:
- **Status Bar**: `👁 Watching 3 PQ files` (when files are being watched)
- **Notifications**: `📝 File changed, syncing: filename.m`
- **Verbose Logs**: Real-time sync details in Output panel

## 🛡️ Backup & Safety Features

### **Automatic Backups**:
- Created before every sync operation
- Timestamped: `filename.xlsx.backup.2025-06-20T18-10-19-087Z`
- Configurable location: same folder, temp, or custom path

### **Backup Management**:
- **Max Backups**: Keep only N most recent (default: 5)
- **Auto-Cleanup**: Delete old backups automatically
- **Manual Cleanup**: Right-click Excel file → "Cleanup Old Backups"

### **Custom Backup Locations**:
```json
// Same folder as Excel file (default)
"excel-power-query-editor.backupLocation": "sameFolder"

// OS temp directory
"excel-power-query-editor.backupLocation": "tempFolder"

// Custom path (relative or absolute)
"excel-power-query-editor.backupLocation": "custom"
"excel-power-query-editor.customBackupPath": "./excel-backups"
```

## 📂 File Structure Examples

### **Simple Workspace**:
```
project/
├── SalesReport.xlsx
├── SalesReport.xlsx_PowerQuery.m           ← Auto-syncs to Excel
└── SalesReport.xlsx.backup.2025-06-20T...  ← Automatic backup
```

### **Multi-File Project**:
```
analytics/
├── Q1_Report.xlsx
├── Q1_Report.xlsx_PowerQuery.m              ← Watching ✓
├── Q2_Report.xlsm  
├── Q2_Report.xlsm_PowerQuery.m              ← Watching ✓
├── Dashboard.xlsb
├── Dashboard.xlsb_PowerQuery.m              ← Watching ✓
└── excel-backups/                           ← Custom backup location
    ├── Q1_Report.xlsx.backup.2025...
    ├── Q2_Report.xlsm.backup.2025...
    └── Dashboard.xlsb.backup.2025...
```

### **Status Bar Display**:
```
👁 Watching 3 PQ files    [Bottom-right corner when files are being watched]
```

## 🔧 Troubleshooting

### **Sync Asks for File Selection**
**Problem**: Sync prompts to select Excel file instead of auto-detecting
**Solutions**:
1. ✅ Check Excel file exists in same directory
2. ✅ Verify naming: `filename.xlsx_PowerQuery.m` format
3. ✅ Ensure Excel extension is `.xlsx`, `.xlsm`, or `.xlsb`
4. ✅ Try placing both files in same folder

### **No Power Query Found**
**Problem**: Extraction reports "No Power Query found"
**Solutions**:
1. ✅ Use **"Raw Excel Extraction"** to see all content
2. ✅ Check if Excel uses external connections instead of Power Query
3. ✅ Verify file contains actual Power Query (Data → Get Data)
4. ✅ Try with known Power Query-enabled file first

### **Auto-Watch Not Working After Reload**
**Problem**: Watch stops after VS Code reload
**Solutions**:
1. ✅ Enable **"Watch Always"** setting for automatic restoration
2. ✅ Check **Verbose Mode** for initialization messages
3. ✅ Manually restart: Right-click `.m` file → "Toggle Watch"
4. ✅ Verify settings: Search "Excel Power Query" in Settings

### **Backup Files Accumulating**
**Problem**: Too many backup files in directory
**Solutions**:
1. ✅ Adjust **"Max Backups"** setting (default: 5)
2. ✅ Use **"Cleanup Old Backups"** command on Excel files
3. ✅ Set custom backup location: `./backups` or temp folder
4. ✅ Disable backups entirely (not recommended): `"autoBackupBeforeSync": false`

### **Slow Performance**
**Problem**: Extension feels sluggish
**Solutions**:
1. ✅ Reduce **"Max Backups"** to 3
2. ✅ Disable **"Verbose Mode"** if not needed
3. ✅ Use temp folder for backups instead of custom path
4. ✅ Consider disabling auto-watch for large projects

## 🐛 Debug Features

### **Raw Extraction** (Advanced)
Access all Excel components for debugging:
1. Right-click Excel file
2. Select **"Raw Excel Extraction (Debug)"**
3. Creates `debug_extraction/` folder with:
   - All XML files from Excel archive
   - Power Query DataMashup content
   - Parsed structure files

### **Verbose Logging**
Enable detailed operation logging:
1. **Settings**: `"excel-power-query-editor.verboseMode": true`
2. **View Output**: `View` → `Output` → "Excel Power Query Editor"
3. **See Logs**: Real-time sync, watch, and backup operations

### **Debug Mode**
Enable enhanced debugging:
1. **Settings**: `"excel-power-query-editor.debugMode": true`
2. **Creates**: Additional debug files during sync operations
3. **Helps With**: Troubleshooting sync failures and Excel format issues

## 📋 Supported File Types

### **Excel Files** (Source)
| Extension | Description | Support Level |
|-----------|-------------|---------------|
| `.xlsx` | Excel Workbook | ✅ Full Support |
| `.xlsm` | Excel Macro-Enabled Workbook | ✅ Full Support |
| `.xlsb` | Excel Binary Workbook | ✅ Full Support |

### **Power Query Files** (Generated)
| Extension | Description | Features |
|-----------|-------------|----------|
| `.m` | Power Query M Language | ✅ Syntax highlighting<br>✅ Auto-sync<br>✅ Comment preservation |

### **Power Query Storage Formats**
| Format | Description | Extraction |
|--------|-------------|------------|
| **DataMashup** | Modern Power Query storage | ✅ Full support with comments |
| **QueryTable** | Legacy query storage | ⚠️ Limited support |
| **Connection** | External data connections | ⚠️ Partial support |

## ⚙️ Advanced Settings Configuration

### **Quick Access**
`File` → `Preferences` → `Settings` → Search "Excel Power Query"

### **Essential Settings**

#### **Auto-Watch & Productivity**
```json
{
  // Auto-watch when extracting files
  "excel-power-query-editor.watchAlways": true,
  
  // Show detailed logs for debugging
  "excel-power-query-editor.verboseMode": true,
  
  // Display watch count in status bar
  "excel-power-query-editor.showStatusBarInfo": true,
  
  // Stop watching when files are deleted
  "excel-power-query-editor.watchOffOnDelete": true
}
```

#### **Backup & Safety**
```json
{
  // Create backups before sync (recommended)
  "excel-power-query-editor.autoBackupBeforeSync": true,
  
  // Custom backup location
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./PQ-backups",
  
  // Keep 5 most recent backups
  "excel-power-query-editor.maxBackups": 5,
  
  // Auto-delete old backups
  "excel-power-query-editor.autoCleanupBackups": true
}
```

#### **User Experience**
```json
{
  // Confirm before sync & delete
  "excel-power-query-editor.syncDeleteAlwaysConfirm": true,
  
  // Stop watching when using Sync & Delete
  "excel-power-query-editor.syncDeleteTurnsWatchOff": true,
  
  // Operation timeout (30 seconds)
  "excel-power-query-editor.syncTimeout": 30000
}
```

### **Recommended Configurations**

#### **🚀 Active Development Setup**
```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.maxBackups": 10,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": false,
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./PQ-backups"
}
```

#### **🛡️ Production/Shared Files Setup**
```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.maxBackups": 3,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": true,
  "excel-power-query-editor.verboseMode": false,
  "excel-power-query-editor.backupLocation": "tempFolder"
}
```

#### **⚡ Performance/Minimal Setup**
```json
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.showStatusBarInfo": false,
  "excel-power-query-editor.verboseMode": false,
  "excel-power-query-editor.watchAlways": false
}
```

### **Settings Scope**

#### **User Settings** (`settings.json`)
Apply to all VS Code workspaces globally.

#### **Workspace Settings** (`.vscode/settings.json`)
Apply only to current project. Example for Power Query development:
```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.customBackupPath": "./backups",
  "excel-power-query-editor.maxBackups": 15
}
```

## 🔍 Monitoring & Debugging

### **Verbose Output Usage**
1. **Enable**: `"excel-power-query-editor.verboseMode": true`
2. **Access**: `View` → `Output` → Select "Excel Power Query Editor"
3. **Monitor**: Real-time logs of all operations:
   ```
   [2025-06-20T18:10:19.087Z] Started watching: SalesReport.xlsx_PowerQuery.m
   [2025-06-20T18:10:25.123Z] File changed, auto-syncing: SalesReport.xlsx_PowerQuery.m
   [2025-06-20T18:10:25.156Z] Backup created: ./backups/SalesReport.xlsx.backup.2025...
   [2025-06-20T18:10:25.234Z] Sync completed successfully
   ```

### **Debug Mode Features**
When `"debugMode": true`:
- 🔍 **Enhanced Error Messages**: Detailed failure analysis
- 📁 **Debug File Creation**: XML structure saved to `debug_sync/` folder
- 🔬 **Raw Content Analysis**: Full Excel content extraction for troubleshooting
- 📊 **Sync Attempt Logging**: Step-by-step sync process details

## ⚠️ Current Limitations

### **Technical Constraints**
- ✅ **No Excel Installation Required** (unlike legacy extensions)
- ✅ **Cross-Platform Support** (Windows, macOS, Linux)
- ✅ **No COM Dependencies** (reliable across VS Code updates)
- ⚠️ **Single Power Query per Excel File** (current implementation)
- ⚠️ **Limited QueryTable Support** (legacy format)

### **File Operation Requirements**
- 📄 **Excel Files Should Be Closed** during sync operations
- 🔒 **Avoid Network Drive Issues** by using local files when possible
- 💾 **Backup Files Created** automatically (can be disabled)

### **Performance Considerations**
- 🚀 **Fast Extraction**: Direct file parsing, no COM overhead
- ⚡ **Quick Sync**: Efficient binary blob updates
- 📊 **Scalable**: Tested with files up to several MB
- 🔄 **Auto-Watch Limit**: Maximum 20 files auto-watched on startup

## 💡 Pro Tips & Best Practices

### **Workflow Optimization**
1. 🎯 **Enable Watch Always** for active Power Query development
2. 📁 **Use Custom Backup Path** like `./PQ-backups` for organization
3. 🔍 **Enable Verbose Mode** during initial setup for visibility
4. ⚡ **Use Toggle Watch** command for quick enable/disable

### **File Management**
1. 📝 **Keep Descriptive Names**: `Q4_Sales_Analysis.xlsx` instead of `report.xlsx`
2. 📂 **Organize by Project**: Separate folders for different analyses
3. 🗂️ **Use Workspace Settings** for project-specific configurations
4. 🔄 **Regular Cleanup**: Use "Cleanup Old Backups" periodically

### **Safety Practices**
1. 🛡️ **Test on Copies** before working on important files
2. 💾 **Verify Backups** are being created in expected location
3. 🔍 **Check Verbose Logs** if operations seem unsuccessful
4. 📊 **Use Debug Mode** for troubleshooting complex sync issues

### **Collaboration**
1. 👥 **Share Workspace Settings** via `.vscode/settings.json` in repository
2. 📁 **Use Relative Backup Paths** like `./backups` for portability
3. 🔄 **Document Watch Status** in project README
4. ⚙️ **Standardize Team Settings** for consistent behavior

## 🤝 Integrations & Credits

### **Core Dependencies**
- **[excel-datamashup](https://github.com/Vladinator/excel-datamashup)** by Vladinator (GPL-3.0)
  - Powers reliable Power Query extraction and sync
  - Handles Excel DataMashup XML parsing and generation
- **[Chokidar](https://github.com/paulmillr/chokidar)** - Robust file watching
- **[JSZip](https://github.com/Stuk/jszip)** - Excel file parsing

### **Recommended Companion Extensions**

```vscode-extensions
powerquery.vscode-powerquery,grapecity.gc-excelviewer
```

- **[Power Query / M Language](https://marketplace.visualstudio.com/items?itemName=powerquery.vscode-powerquery)** *(Required)*
  - Essential for M language syntax highlighting and IntelliSense
  - Automatically installed via Extension Pack
  - Provides proper code completion and error detection
- **[Excel Viewer by GrapeCity](https://marketplace.visualstudio.com/items?itemName=grapecity.gc-excelviewer)** *(Optional)*
  - View Excel files directly in VS Code without opening Excel
  - Perfect companion for Power Query development workflow
  - Seamless integration with this extension

### **Version History**
- **v0.4.x**: Extension Pack with Power Query M Language, improved categories and documentation
- **v0.4.1**: Auto-watch initialization, hybrid activation
- **v0.4.0**: Backup management, cleanup commands
- **v0.3.1**: Settings implementation, auto-watch fixes
- **v0.2.2**: Sync improvements, binary blob handling
- **v0.1.3**: Initial stable release

---

## 📞 Support & Feedback

### **💖 Support This Project**
If this extension makes your Power Query development more productive, consider supporting its continued development:

[![Buy Me a Coffee](https://img.shields.io/badge/-Buy%20Me%20a%20Coffee-ffdd00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/ewc3labs)

*Your support helps maintain and improve this extension for the entire Power Query community!*

### **Getting Help**
1. 🔍 **Check Verbose Logs**: Enable verbose mode for detailed operation info
2. 🐛 **Use Debug Mode**: For complex sync issues
3. 🔧 **Try Raw Extraction**: For troubleshooting extraction problems
4. 📖 **Consult Settings**: Many behaviors are configurable

### **Known Working Configurations**
- ✅ **Windows 11** with Excel 2021 (.xlsx, .xlsm, .xlsb)
- ✅ **Cross-platform** VS Code (Windows, macOS, Linux)
- ✅ **Large Files** up to several MB with complex Power Query
- ✅ **Network Drives** (with proper permissions)

*This extension provides a modern, reliable alternative to COM-based Power Query editing solutions.*

---

**📝 Last Updated**: June 2025  
**📄 For installation and overview**: See `README.md`  
**⚙️ For quick settings**: See `CONFIGURATION.md`
