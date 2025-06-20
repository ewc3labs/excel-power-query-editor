# Excel Power Query Editor

> **A modern, reliable VS Code extension for editing Power Query M code from Excel files**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VS Code Marketplace](https://img.shields.io/badge/VS%20Code-Marketplace-blue.svg)](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor)
[![Buy Me a Coffee](https://img.shields.io/badge/-Buy%20Me%20a%20Coffee-ffdd00?style=flat-square&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/ewc3labs)

## � Installation

### **From VS Code Marketplace (Recommended)**

1. **VS Code Extensions View**: 
   - Open VS Code → Extensions (Ctrl+Shift+X)
   - Search for "Excel Power Query Editor" 
   - Click Install

2. **Command Line**:
   ```bash
   code --install-extension ewc3labs.excel-power-query-editor
   ```

3. **Direct Link**: [Install from Marketplace](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor)

### **Alternative: From VSIX File**
Download and install a specific version manually:
```bash
code --install-extension excel-power-query-editor-[version].vsix
```

## 🚨 IMPORTANT: Required Extension

**This extension requires the Microsoft Power Query / M Language extension for proper syntax highlighting and IntelliSense:**

```vscode-extensions
powerquery.vscode-powerquery
```

*The Power Query extension will be automatically installed when you install this extension (via Extension Pack).*

## 📚 Complete Documentation

- **📖 [Complete User Guide](USER_GUIDE.md)** - Detailed usage instructions, features, and troubleshooting
- **⚙️ [Configuration Guide](CONFIGURATION.md)** - Quick reference for all settings 
- **📝 [Changelog](CHANGELOG.md)** - Version history and updates

## ⚡ Quick Start

1. **Install**: Search "Excel Power Query Editor" in Extensions view
2. **Open Excel file**: Right-click `.xlsx`/`.xlsm` → "Extract Power Query from Excel"  
3. **Edit**: Modify the generated `.m` file with full VS Code features
4. **Auto-Sync**: Right-click `.m` file → "Toggle Watch" for automatic sync on save
5. **Enjoy**: Modern Power Query development workflow! 🎉

## Why This Extension?

Excel's Power Query editor is **painful to use**. This extension brings the **power of VS Code** to Power Query development:

- 🚀 **Modern Architecture**: No COM/ActiveX dependencies that break with VS Code updates
- 🔧 **Reliable**: Direct Excel file parsing - no Excel installation required  
- 🌐 **Cross-Platform**: Works on Windows, macOS, and Linux
- ⚡ **Fast**: Instant startup, no waiting for COM objects
- 🎨 **Beautiful**: Syntax highlighting, IntelliSense, and proper formatting

## The Problem This Solves

**Original EditExcelPQM extension** (and Excel's built-in editor) suffer from:
- ❌ Breaks with every VS Code update (COM/ActiveX issues)
- ❌ Windows-only, requires Excel installed
- ❌ Leaves Excel zombie processes  
- ❌ Unreliable startup (popup dependencies)
- ❌ Terrible editing experience

**This extension** provides:
- ✅ Update-resistant architecture
- ✅ Works without Excel installed
- ✅ Clean, reliable operation
- ✅ Cross-platform compatibility
- ✅ Modern VS Code integration

## Features

- **Extract Power Query from Excel**: Right-click on `.xlsx` or `.xlsm` files to extract Power Query definitions to `.m` files
- **Edit with Syntax Highlighting**: Full Power Query M language support with syntax highlighting
- **Auto-Sync**: Watch `.m` files for changes and automatically sync back to Excel
- **No COM Dependencies**: Works without Excel installed, uses direct file parsing
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Usage

### Extract Power Query from Excel

1. Right-click on an Excel file (`.xlsx` or `.xlsm`) in the Explorer
2. Select "Extract Power Query from Excel"
3. The extension will create `.m` files in a new folder next to your Excel file
4. Open the `.m` files to edit your Power Query code

### Edit Power Query Code

- `.m` files have full syntax highlighting for Power Query M language
- IntelliSense support for Power Query functions and keywords
- Proper indentation and bracket matching

### Sync Changes Back to Excel

1. Open a `.m` file
2. Right-click in the editor and select "Sync Power Query to Excel"
3. Or use the sync button in the editor toolbar
4. The extension will update the corresponding Excel file

### Auto-Watch for Changes

1. Open a `.m` file
2. Right-click and select "Watch Power Query File"
3. The extension will automatically sync changes to Excel when you save
4. A status bar indicator shows the watching status

## Commands

- `Excel Power Query: Extract from Excel` - Extract Power Query definitions from Excel file (creates `filename_PowerQuery.m` in same folder)
- `Excel Power Query: Sync to Excel` - Sync current .m file back to Excel
- `Excel Power Query: Sync & Delete` - Sync .m file to Excel and delete the .m file (with confirmation)
- `Excel Power Query: Watch File` - Start watching current .m file for automatic sync on save
- `Excel Power Query: Stop Watching` - Stop watching current file
- `Excel Power Query: Raw Extraction (Debug)` - Extract all Excel content for debugging

## Requirements

- VS Code 1.96.0 or later
- No Excel installation required (uses direct file parsing)

## Known Limitations

- Currently supports basic Power Query extraction (advanced features coming soon)
- Excel file backup is created automatically before modifications
- Some complex Power Query features may not be fully supported yet

## Development

This extension is built with:
- TypeScript
- xlsx library for Excel file parsing
- chokidar for file watching
- esbuild for bundling

### Building from Source

```bash
npm install
npm run compile
```

### Testing

```bash
npm test
```

## Acknowledgments

Inspired by the original [EditExcelPQM](https://github.com/amalanov/EditExcelPQM) by Alexander Malanov, but completely rewritten with modern architecture to solve reliability issues.

## ⚙️ Settings

The extension provides comprehensive settings for customizing your workflow. Access via `File` > `Preferences` > `Settings` > search "Excel Power Query":

### **Watch & Auto-Sync Settings**

| Setting | Default | Description |
|---------|---------|-------------|
| **Watch Always** | `false` | Automatically start watching when extracting Power Query files. Perfect for active development. |
| **Watch Off On Delete** | `true` | Automatically stop watching when .m files are deleted (prevents zombie watchers). |
| **Sync Delete Turns Watch Off** | `true` | Stop watching when using "Sync & Delete" command. |
| **Show Status Bar Info** | `true` | Display watch status in status bar (e.g., "👁 Watching 3 PQ files"). |

### **Backup & Safety Settings**

| Setting | Default | Description |
|---------|---------|-------------|
| **Auto Backup Before Sync** | `true` | Create automatic backups before syncing to Excel files. |
| **Backup Location** | `"sameFolder"` | Where to store backup files: `"sameFolder"`, `"tempFolder"`, or `"custom"`. |
| **Custom Backup Path** | `""` | Custom path for backups (when Backup Location is "custom"). Supports relative paths like `./backups`. |
| **Max Backups** | `5` | Maximum backup files to keep per Excel file (1-50). Older backups are auto-deleted. |
| **Auto Cleanup Backups** | `true` | Automatically delete old backups when exceeding Max Backups limit. |

### **User Experience Settings**

| Setting | Default | Description |
|---------|---------|-------------|
| **Sync Delete Always Confirm** | `true` | Ask for confirmation before "Sync & Delete" (uncheck for instant deletion). |
| **Verbose Mode** | `false` | Show detailed logging in Output panel for debugging and monitoring. |
| **Debug Mode** | `false` | Enable advanced debug logging and save debug files for troubleshooting. |
| **Sync Timeout** | `30000` | Timeout in milliseconds for sync operations (5000-120000). |

### **Example Workflows**

**🔄 Active Development Setup:**
```json
{
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.maxBackups": 10
}
```

**🛡️ Conservative/Production Setup:**
```json
{
  "excel-power-query-editor.watchAlways": false,
  "excel-power-query-editor.maxBackups": 3,
  "excel-power-query-editor.backupLocation": "custom",
  "excel-power-query-editor.customBackupPath": "./excel-backups"
}
```

**⚡ Speed/Minimal Setup:**
```json
{
  "excel-power-query-editor.autoBackupBeforeSync": false,
  "excel-power-query-editor.syncDeleteAlwaysConfirm": false,
  "excel-power-query-editor.showStatusBarInfo": false
}
```

### **Accessing Verbose Output**

When Verbose Mode is enabled:
1. Go to `View` > `Output`
2. Select "Excel Power Query Editor" from the dropdown
3. See detailed logs of all operations, watch events, and errors

## 💖 Support This Project

If this extension saves you time and makes your Power Query development more enjoyable, consider supporting its development:

[![Buy Me a Coffee](https://img.shields.io/badge/-Buy%20Me%20a%20Coffee-ffdd00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/ewc3labs)

Your support helps:
- 🛠️ **Continue development** and add new features
- 🐛 **Fix bugs** and improve reliability
- 📚 **Maintain documentation** and user guides
- 💡 **Respond to feature requests** from the community

*Even a small contribution makes a big difference!*

## Contributing

Contributions are welcome! This extension is built to serve the Power Query community.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/ewc3/excel-power-query-editor/blob/HEAD/LICENSE) file for details.

---

**Made with ❤️ for the Power Query community by [EWC3 Labs](https://github.com/ewc3)**

*Because editing Power Query in Excel shouldn't be painful.*

---

**☕ Enjoying this extension?** [Buy me a coffee](https://www.buymeacoffee.com/ewc3labs) to support continued development!

## Credits and Attribution

This extension uses the excellent [excel-datamashup](https://github.com/Vladinator/excel-datamashup) library by [Vladinator](https://github.com/Vladinator) for robust Excel Power Query extraction. The excel-datamashup library is licensed under GPL-3.0 and provides the core functionality for parsing Excel DataMashup binary formats.

**Special thanks to:**
- **[Vladinator](https://github.com/Vladinator)** for creating the excel-datamashup library that makes reliable Power Query extraction possible
- The Power Query community for feedback and inspiration

This VS Code extension adds the user interface, file management, and editing workflow on top of the excel-datamashup parsing engine.

## 🤝 Recommended Extensions

This extension works best with these companion extensions:

```vscode-extensions
powerquery.vscode-powerquery,grapecity.gc-excelviewer
```

- **[Power Query / M Language](https://marketplace.visualstudio.com/items?itemName=powerquery.vscode-powerquery)** *(Required)* - Provides syntax highlighting and IntelliSense for .m files
- **[Excel Viewer by GrapeCity](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer)** *(Optional)* - View Excel files directly in VS Code for seamless workflow

*The Power Query extension is automatically installed via Extension Pack when you install this extension.*
