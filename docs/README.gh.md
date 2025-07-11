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

<!-- BADGES -->
<p align="center">
  <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-yellow.svg">
  <img alt="CI/CD" src="https://img.shields.io/badge/CI%2FCD%20Pipeline-failing-red.svg">
  <img alt="Tests Passing" src="https://img.shields.io/badge/tests-63%20passing-brightgreen.svg">
  <img alt="VS Code" src="https://img.shields.io/badge/VS_Code-Marketplace-blue.svg">
  <img alt="Buy Me a Coffee" src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-yellow?logo=buy-me-a-coffee&logoColor=white">
</p>
<!-- /BADGES -->

---

### 🛠️ About This Extension

At **EWC3 Labs**, we don’t just build tools — we rage-build solutions to common problems that grind our gears on the daily. We got tired of fighting Excel’s half-baked Power Query editor and decided to _**just rip the M code**_ straight into VS Code, where it belongs and where CoPilot _lives_. Other devs built the foundational pieces _(see Acknowledgments below)_, and we stitched them together like caffeinated mad scientists in a lightning storm.

This extension exists because the existing workflow is clunky, fragile, and dumb. There’s no Excel or COM (_or Windows_) requirement, and no popup that says “something went wrong” with no actionable info. Just clean `.m` files. One context. Full references. You save — we sync. Done.

This is Dev/Power User tooling that finally respects your time.

---

## ⚡ Quick Start

### 1. Install

Open VS Code → Extensions (`Ctrl+Shift+X`) → Search **"Excel Power Query Editor"** → Install

### 2. Extract & Edit

1. Right-click any Excel file (`.xlsx`, `.xlsm`, `.xlsb`) in Explorer
2. Select **"Extract Power Query from Excel"**
3. Edit the generated `.m` file with full VS Code features

### 3. Auto-Sync

1. Right-click the `.m` file → **"Toggle Watch"**
2. Your changes automatically sync to Excel when you save
3. Automatic backups keep your data safe

## 🚀 Key Features

- **🔄 Bidirectional Sync**: Extract from Excel → Edit in VS Code → Sync back seamlessly
- **👁️ Auto-Watch Mode**: Real-time sync when you save (with intelligent debouncing)
- **🛡️ Smart Backups**: Automatic Excel backups before any changes
- **🔧 Zero Dependencies**: No Excel installation required, works on Windows/Mac/Linux
- **💡 Full IntelliSense**: Complete M language support with syntax highlighting
- **⚙️ Highly Configurable**: Customize backup locations, watch behavior, sync timing

## Why This Extension?

Excel's Power Query editor is **painful to use**. This extension brings the **power of VS Code** to Power Query development:

- 🚀 **Modern Architecture**: No COM/ActiveX dependencies that break with VS Code updates
- 🔧 **Reliable**: Direct Excel file parsing - no Excel installation required
- 🌐 **Cross-Platform**: Works on Windows, macOS, and Linux
- ⚡ **Fast**: Instant startup, no waiting for COM objects
- 🎨 **Beautiful**: Syntax highlighting, IntelliSense, and proper formatting

## The Problem This Solves

**Excel's built-in editor** and legacy extensions suffer from:

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

## 📚 Complete Documentation

- **📖 [User Guide](docs/USER_GUIDE.md)** - Complete workflows, advanced features, troubleshooting
- **⚙️ [Configuration](docs/CONFIGURATION.md)** - All settings, examples, use cases
- **🤝 [Contributing](docs/CONTRIBUTING.md)** - Development setup, testing, contribution guidelines
- **📝 [Changelog](CHANGELOG.md)** - Version history and feature updates

## 🆘 Need Help?

- **Issues**: [GitHub Issues](https://github.com/ewc3labs/excel-power-query-editor/issues)
- **Discussions**: [GitHub Discussions](https://github.com/ewc3labs/excel-power-query-editor/discussions)
- **Support**: [![Buy Me a Coffee](https://img.shields.io/badge/-Buy%20Me%20a%20Coffee-ffdd00?style=flat&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/ewc3labs)

## 🤝 Acknowledgments & Credits

This extension builds upon the excellent work of several key contributors to the Power Query ecosystem:

**Inspired by:**

- **[Alexander Malanov](https://github.com/amalanov)** - Creator of the original [EditExcelPQM](https://github.com/amalanov/EditExcelPQM) extension, which pioneered Power Query editing in VS Code

**Powered by:**

- **[Microsoft Power Query / M Language Extension](https://marketplace.visualstudio.com/items?itemName=powerquery.vscode-powerquery)** - Provides essential M language syntax highlighting and IntelliSense
- **[MESCIUS Excel Viewer](https://marketplace.visualstudio.com/items?itemName=MESCIUS.gc-excelviewer)** - Enables Excel file viewing in VS Code for seamless CoPilot workflows

**Technical Foundation:**

- **[excel-datamashup](https://github.com/Vladinator/excel-datamashup)** by [Vladinator](https://github.com/Vladinator) - Robust Excel Power Query extraction library

This extension represents a complete architectural rewrite focused on reliability, cross-platform compatibility, and modern VS Code integration patterns.

---

**Excel Power Query Editor** - _Because Power Query development shouldn't be painful_ ✨
