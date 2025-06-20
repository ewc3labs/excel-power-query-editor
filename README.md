# Excel Power Query Editor

> **A modern, reliable VS Code extension for editing Power Query M code from Excel files**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VS Code Marketplace](https://img.shields.io/badge/VS%20Code-Marketplace-blue.svg)](https://marketplace.visualstudio.com/items?itemName=ewc3labs.excel-power-query-editor)

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

## Credits and Attribution

This extension uses the excellent [excel-datamashup](https://github.com/Vladinator/excel-datamashup) library by [Vladinator](https://github.com/Vladinator) for robust Excel Power Query extraction. The excel-datamashup library is licensed under GPL-3.0 and provides the core functionality for parsing Excel DataMashup binary formats.

**Special thanks to:**
- **[Vladinator](https://github.com/Vladinator)** for creating the excel-datamashup library that makes reliable Power Query extraction possible
- The Power Query community for feedback and inspiration

This VS Code extension adds the user interface, file management, and editing workflow on top of the excel-datamashup parsing engine.

## Following extension guidelines

Ensure that you've read through the extensions guidelines and follow the best practices for creating your extension.

* [Extension Guidelines](https://code.visualstudio.com/api/references/extension-guidelines)

## Working with Markdown

You can author your README using Visual Studio Code. Here are some useful editor keyboard shortcuts:

* Split the editor (`Cmd+\` on macOS or `Ctrl+\` on Windows and Linux).
* Toggle preview (`Shift+Cmd+V` on macOS or `Shift+Ctrl+V` on Windows and Linux).
* Press `Ctrl+Space` (Windows, Linux, macOS) to see a list of Markdown snippets.

## For more information

* [Visual Studio Code's Markdown Support](http://code.visualstudio.com/docs/languages/markdown)
* [Markdown Syntax Reference](https://help.github.com/articles/markdown-basics/)

**Enjoy!**

## 🤝 **Recommended Extensions**

This extension works great with **[Excel Viewer by GrapeCity](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer)** which displays Excel files natively in VS Code. Together they provide:

- **View Excel files** directly in VS Code (supports .xlsx, .xlsm)
- **Edit Power Query** in the same workspace
- **Seamless workflow** without leaving VS Code

*The Excel Viewer will be automatically suggested when you install this extension.*
