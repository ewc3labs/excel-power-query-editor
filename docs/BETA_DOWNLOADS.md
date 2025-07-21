# Excel Power Query Editor
A modern, reliable VS Code extension for editing Power Query M code directly from Excel files

---

# 🧪 Beta Downloads & Nightly Builds

Get early access to the latest features and fixes before they hit the VS Code Marketplace!

## 🚀 Quick Install

[![Latest Pre-release](https://img.shields.io/github/v/release/ewc3labs/excel-power-query-editor?include_prereleases&label=latest%20beta)](https://github.com/ewc3labs/excel-power-query-editor/releases)

**Reliable method:**
1. Go to [Releases](https://github.com/ewc3labs/excel-power-query-editor/releases)
2. Download the latest `.vsix` file from a "Pre-release" entry
3. Install: `code --install-extension excel-power-query-editor-*.vsix`

## 🔄 Auto-Update Script

Save this as `update-excel-pq-beta.sh` (or `.bat` for Windows):

```bash
#!/bin/bash
# Download and install latest Excel Power Query Editor beta

echo "🔍 Fetching latest beta release..."
LATEST_URL=$(curl -s https://api.github.com/repos/ewc3labs/excel-power-query-editor/releases | jq -r '.[0].assets[0].browser_download_url')

if [[ "$LATEST_URL" != "null" ]]; then
    echo "📦 Downloading: $LATEST_URL"
    curl -L -o excel-power-query-editor-beta.vsix "$LATEST_URL"
    
    echo "🚀 Installing..."
    code --install-extension excel-power-query-editor-beta.vsix
    
    echo "✅ Beta installed! Restart VS Code to use."
    rm excel-power-query-editor-beta.vsix
else
    echo "❌ Could not fetch latest release"
fi
```

## 📋 What's in Beta?

Beta releases include:
- 🆕 **New Features** - Latest functionality before marketplace release
- 🐛 **Bug Fixes** - Immediate fixes for reported issues  
- ⚡ **Performance Improvements** - Speed and reliability enhancements
- 🧪 **Experimental Features** - Try cutting-edge capabilities

## ⚠️ Beta Considerations

- **Stability:** Generally stable, but may have occasional issues
- **Feedback:** Please [report any bugs](https://github.com/ewc3labs/excel-power-query-editor/issues/new) you find!
- **Updates:** New betas released automatically when code is pushed
- **Rollback:** Keep stable version handy in case you need to revert

## 🔗 Beta Release Channels

- **🏷️ Release Candidates (RC):** `v0.5.0-rc.1`, `v0.5.0-rc.2` - Near-final versions
- **🌙 Nightly Builds:** Automatic builds from latest `release/` branch commits
- **🔥 Hotfixes:** Critical fixes released immediately as needed

## 📞 Support

Having issues with a beta? 
- [📋 Check existing issues](https://github.com/ewc3labs/excel-power-query-editor/issues)
- [🆕 Report new bugs](https://github.com/ewc3labs/excel-power-query-editor/issues/new)
- [💬 Discussion forum](https://github.com/ewc3labs/excel-power-query-editor/discussions)

---

<p align="center">
  <img src="assets/EWC3LabsLogo-blue-128x128.png" width="128" height="128" alt="Georgie the QA Officer"><br>
  <sub><b>Georgie, our QA Officer</b></sub>
</p>

---

**Happy testing!** 🧪✨
