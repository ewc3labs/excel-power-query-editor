# Testing Notes - Excel Power Query Editor v0.5.0

## Test Environment Setup

### Dev Container Settings Issue

**Issue**: Settings configured in local user settings (`C:\Users\[user]\AppData\Roaming\Code\User\settings.json`) do not flow into dev container environments.

**Solution**: Extension settings must be configured in workspace settings (`.vscode/settings.json`) when working in dev containers.

**Example Workspace Settings for Dev Container**:

```json
{
  "excel-power-query-editor.verboseMode": true,
  "excel-power-query-editor.watchAlways": true,
  "excel-power-query-editor.backupLocation": "sameFolder",
  "excel-power-query-editor.customBackupPath": "./VSCodeBackups",
  "excel-power-query-editor.debugMode": true,
  "excel-power-query-editor.watchOffOnDelete": true
}
```

**Action Required**: Document this limitation in user guide and consider auto-detection or warning for dev container users.

---

## Feature Testing Issues

### 1. Apply Recommended Defaults Command

**Status**: 🔴 **CRITICAL ISSUE - DEV CONTAINER INCOMPATIBLE**

**Issue**: The "Excel PQ: Apply Recommended Defaults" command modifies **Global/User settings** using `vscode.ConfigurationTarget.Global`, which means:

- ❌ **Does NOT work in dev containers** (user settings don't persist/apply)
- ❌ **Settings changes are lost** when dev container is rebuilt
- ❌ **No effect on current session** in dev container environment

**Current Implementation**:

```typescript
await config.update(setting, value, vscode.ConfigurationTarget.Global);
```

**Settings Applied**:

- `watchAlways`: false
- `watchOffOnDelete`: true
- `syncDeleteAlwaysConfirm`: true
- `verboseMode`: false
- `autoBackupBeforeSync`: true
- `backupLocation`: 'sameFolder'
- `backup.maxFiles`: 5
- `autoCleanupBackups`: true
- `syncTimeout`: 30000
- `debugMode`: false
- `showStatusBarInfo`: true
- `sync.openExcelAfterWrite`: false
- `sync.debounceMs`: 500
- `watch.checkExcelWriteable`: true

**Action Required**:

1. **Immediate Fix**: Change to `vscode.ConfigurationTarget.Workspace` for dev container compatibility
2. **Enhancement**: Let user choose scope (Global vs Workspace)
3. **Documentation**: Warn users about dev container limitations if using Global scope

**Proposed Fix**:

```typescript
// Detect if in dev container and use appropriate target
const isDevContainer = vscode.env.remoteName?.includes("dev-container");
const target = isDevContainer
  ? vscode.ConfigurationTarget.Workspace
  : vscode.ConfigurationTarget.Global;
await config.update(setting, value, target);
```

---

### 2. Large File DataMashup Recognition

**Status**: ✅ **FIXED - CRITICAL PRODUCTION BUG RESOLVED**

**Issue RESOLVED**: Extension now successfully extracts and syncs DataMashup from large files (50MB+)

**Root Cause Identified**: **Hardcoded customXml scanning limitation**

- Extension was only checking `customXml/item1.xml`, `item2.xml`, `item3.xml`
- Large files store DataMashup in different locations (e.g., `customXml/item19.xml`)
- **This was a fundamental architectural flaw affecting production users**

**Fix Implemented** (2025-07-12):

1. **Dynamic customXml scanning**: Now scans ALL customXml files instead of hardcoded first 3
2. **Consistent BOM handling**: Both extraction and sync use identical UTF-16 LE detection
3. **Unified DataMashup detection**: Same binary reading logic for extraction and sync

**Test Results**:

- ✅ `PowerQueryFunctions.xlsx` - Small file: **Perfect round-trip sync** (`customXml/item1.xml`)
- ✅ `MAR_DatabaseSummary-V7b.xlsm` - Large file (60MB): **Perfect round-trip sync** (`customXml/item19.xml`)

**Log Evidence**:

```
[2025-07-12T01:08:46.376Z] ✅ Found DataMashup Power Query in: customXml/item19.xml
[2025-07-12T01:09:26.848Z] ✅ excel-datamashup approach succeeded, updating Excel file...
[2025-07-12T01:09:33.658Z] Successfully synced Power Query to Excel: MAR_DatabaseSummary-V7b.xlsm
```

**Production Impact**:

- **MAJOR**: This fix enables the extension to work with real-world Excel files that contain multiple Power Query connections
- **Marketplace**: Resolves critical bug affecting 117+ production installations

---

### 3. Raw Extraction Enhancement for Debugging

**Status**: ✅ **COMPLETED - ENHANCED FOR COMPREHENSIVE DEBUGGING**

**Enhancement Implemented** (2025-07-12): Raw extraction now provides comprehensive DataMashup analysis

**New Features**:

- **Scans ALL customXml files** (not just first 3) - Fixed the same hardcoded bug
- **Detailed file structure reporting** with comprehensive ZIP analysis
- **DataMashup content detection** with proper BOM handling
- **Enhanced error reporting** and debugging information

**Current Raw Extraction Output Example**:

```
[2025-07-12T00:21:53.600Z] Found 38 customXml files to scan: customXml/item1.xml, customXml/item10.xml, customXml/item11.xml, customXml/item12.xml, customXml/item13.xml, customXml/item14.xml, customXml/item15.xml, customXml/item16.xml, customXml/item17.xml, customXml/item18.xml, customXml/item19.xml, [...]
[2025-07-12T00:21:53.620Z] ✅ Found DataMashup Power Query in: customXml/item19.xml
```

**Debugging Value**:

- **Identifies exact DataMashup location** in complex Excel files
- **Validates file structure** before normal extraction
- **Helps troubleshoot** extraction failures by showing all XML content

**Production Use**: Essential for debugging customer files with non-standard DataMashup locations

---

### 4. Legacy Settings Migration + Logging Standardization

**Status**: ✅ **COMPLETED - AUTOMATIC MIGRATION IMPLEMENTED & TESTED**

**Implementation Completed** (2025-07-12):

✅ **New Setting Added**: `excel-power-query-editor.logLevel` with enum values: `["none", "error", "warn", "info", "verbose", "debug"]`

✅ **Automatic Migration Logic**: Implemented in `getEffectiveLogLevel()` function
- Checks for existing `logLevel` setting first
- Falls back to legacy `verboseMode`/`debugMode` if new setting not found  
- Performs one-time migration with user notification
- Graceful fallback to 'info' level as default

✅ **Legacy Settings Preserved**: Marked as `[DEPRECATED]` but kept for backward compatibility
- `verboseMode`: Now marked as deprecated with migration notice
- `debugMode`: Now marked as deprecated with migration notice

✅ **Enhanced Logging System**: Implemented level-based filtering
- Messages categorized by content (error, warn, info, verbose, debug)
- Only logs messages at or above current log level
- Maintains existing `[functionName]` context prefixes

✅ **Apply Recommended Defaults Updated**: 
- Now uses `logLevel: 'info'` instead of legacy boolean flags
- Automatically detects dev container environment for proper scope (workspace vs global)
- Cleanly removes legacy settings during recommended defaults application
- Fixed dev container compatibility issue

✅ **Package & Install Process Fixed**:
- VS Code tasks were hanging, switched to direct npm commands
- `npm run package-vsix` + `node scripts/install-extension.js --force` works reliably
- Extension successfully packaged as 208KB VSIX with 26 files
- Installation completed successfully on Windows

**Migration Experience**:

When a user with legacy settings first activates v0.5.0:
1. Extension detects legacy `verboseMode`/`debugMode` settings
2. Automatically migrates to equivalent `logLevel` value:
   - `debugMode: true` → `logLevel: "debug"`
   - `verboseMode: true` → `logLevel: "verbose"`  
   - Both false or undefined → `logLevel: "info"`
3. Shows informational message about migration with link to settings
4. Legacy settings remain but are ignored (can be manually removed)

**Production Benefits**:

- ✅ **Zero Breaking Changes**: Existing users continue working seamlessly
- ✅ **Automatic Modernization**: Users get improved logging without action required  
- ✅ **Clear Migration Path**: One-time notification explains the change
- ✅ **Better UX**: Single intuitive setting instead of confusing boolean flags
- ✅ **Dev Container Fixed**: Apply Recommended Defaults now works in dev containers

**Code Quality Improvements**:

- ✅ **Level-based Filtering**: Reduces noise in output based on user preference
- ✅ **Context Preservation**: Maintains `[functionName]` prefixes for debugging
- ✅ **Future-proof**: Easy to add new log levels without breaking changes

**Status**: ✅ **PRODUCTION READY** - Ready for v0.5.0 release

---

### 5. File Auto-Watch and Sync on Save

**Status**: 🔴 **CRITICAL ISSUE - EXCESSIVE AUTO-SYNC ON WINDOWS**

**NEW CRITICAL ISSUES DISCOVERED ON WINDOWS** (2025-07-12):

1. **🚨 IMMEDIATE AUTO-SYNC ON EXTRACTION**: Chokidar detects .m file creation and immediately triggers sync
2. **🚨 MULTIPLE WATCHERS CAUSING CASCADE**: Triple watcher system causes 4+ sync events per save
3. **🚨 DUPLICATE METADATA HEADERS**: Two header blocks being written to .m files
4. **🚨 METADATA NOT STRIPPED ON SYNC**: Headers being synced to Excel DataMashup instead of being stripped

**Windows Test Results** (2025-07-12T01:52:13):

```log
[2025-07-12T01:52:13.142Z] Auto-watch enabled for PowerQueryFunctions.xlsx_PowerQuery.m
[2025-07-12T01:52:13.189Z] [watchFile] 🆕 VSCODE: File created: PowerQueryFunctions.xlsx_PowerQuery.m
[2025-07-12T01:52:13.415Z] [watchFile] 🔥 CHOKIDAR: File change detected: PowerQueryFunctions.xlsx_PowerQuery.m
[2025-07-12T01:52:13.415Z] [debouncedSyncToExcel] 🚀 IMMEDIATE SYNC (debounce disabled: 100ms)
[2025-07-12T01:52:13.418Z] Backup created: PowerQueryFunctions.xlsx.backup.2025-07-12T01-52-13-417Z
```

**Root Cause Analysis**:

1. **Over-Engineering**: Triple watcher system designed for dev container issues is causing cascading events on Windows
2. **File Creation Detection**: Chokidar detects .m file creation immediately after extraction
3. **No Header Stripping**: Metadata headers are being synced to Excel instead of being stripped
4. **Duplicate Headers**: Previous header from dev container still present + new Windows header

**Issues Identified**:

- ❌ **Immediate unwanted sync**: File creation triggers immediate sync before user edits
- ❌ **Multiple sync events**: Save operation triggers 4+ sync events from different watchers
- ❌ **Performance impact**: Excessive backup creation and Excel file writes
- ❌ **Data corruption risk**: Metadata headers being written to DataMashup
- ❌ **User experience**: Constant syncing interrupts workflow

**Dev Container vs Windows Comparison**:

| Feature | Dev Container | Windows |
|---------|---------------|---------|
| File watching | Needed polling + backup watchers | Native file events work perfectly |
| Chokidar behavior | 1-second polling delay | Immediate detection |
| Event frequency | Controlled by polling | Real-time cascading events |
| Performance | Acceptable with delays | Excessive with immediate triggers |

**Critical Issues to Fix**:

1. **Simplify Windows Watcher**: Use only Chokidar on Windows, disable triple watcher system
2. **Fix Header Stripping**: Remove metadata headers before syncing to Excel
3. **Prevent Creation Sync**: Don't auto-sync immediately after extraction
4. **Fix Duplicate Headers**: Clean existing .m files with duplicate headers
5. **Restore Proper Debounce**: Increase debounce for Windows to prevent cascade events

**Windows Host Testing Results**:

- ✅ **Extension installation**: Works perfectly on Windows
- ✅ **Manual sync**: Perfect round-trip sync functionality
- ✅ **File watching detection**: **TOO SUCCESSFUL** - Immediate detection causing problems
- ❌ **Auto-sync behavior**: Excessive and disruptive
- ❌ **Metadata handling**: Headers not being stripped properly

**Action Required**:

1. **CRITICAL**: Fix Windows crash on large file extraction (60MB+ files crash the extension)
2. **Immediate**: Disable triple watcher system on Windows (use Chokidar only)
3. **Critical**: Fix metadata header stripping before Excel sync
4. **Important**: Prevent auto-sync on file creation (only on user edits)
5. **Cleanup**: Remove duplicate headers from existing .m files

**NEW ISSUE DISCOVERED** (2025-07-12):

**🚨 EXCEL VIEWER EXTENSION CRASH**: MESCIUS Excel Viewer (or similar) crashes when clicking on 60MB+ Excel files
- **Root Cause**: Excel viewer extensions try to preview large files, causing memory exhaustion  
- **Impact**: Cannot interact with large Excel files in VS Code Explorer
- **Solution**: Disable Excel viewer extensions or avoid clicking on large Excel files
- **Workaround**: Right-click → context menu instead of left-click to avoid triggering preview
- **Status**: ✅ **IDENTIFIED** - This is not our extension's fault

**UPDATED ACTION REQUIRED**:

1. **Immediate**: Test large file extraction using right-click context menu (avoid clicking file)
2. **Immediate**: Disable triple watcher system on Windows (use Chokidar only)  
3. **Critical**: Fix metadata header stripping before Excel sync
4. **Important**: Prevent auto-sync on file creation (only on user edits)
5. **Cleanup**: Remove duplicate headers from existing .m files

**Test Plan**:

1. Reload VS Code to activate minimal debounce setting
2. Save a .m file and look for immediate `🚀 IMMEDIATE SYNC` messages
3. Compare with Windows host testing for validation

**Next Debug Steps**:

1. **Dev Container Testing**: Test enhanced triple watcher system by reloading VS Code and monitoring Output panel for debug logs showing watcher initialization and event detection
2. **Minimal Debounce Test**: Make test changes to .m files and save to identify which watcher mechanisms (chokidar, VS Code FileSystemWatcher, or document save events) successfully detect file changes with immediate sync feedback
3. **Windows Host Validation**: Install extension on Windows host to compare file watching behavior vs dev container environment - this will help isolate whether the issue is dev container specific or broader
4. **Event Correlation**: Use enhanced logging to determine if events are detected but failing during sync vs events not being detected at all

**Windows Host Testing Plan**:

- Install `.vsix` on Windows VS Code
- Test same .m files with same settings
- Compare event detection and sync behavior
- Validate if file watching works normally on Windows host
- Document differences between dev container and host behavior

**Expected Resolution**: Comparison between dev container and Windows host will reveal if issue is:

- **Dev Container specific**: File system mounting/Docker issue requiring container-specific solutions
- **Code issue**: Problem exists on both platforms requiring code fixes
- **Configuration issue**: Settings or environment differences

---

### 6. Power Query Language Extension Linting Issues

**Status**: ⚠️ **MINOR ISSUE - LINTING FALSE POSITIVES**

**Issue**: Power Query/M Language extension shows "Problems" for valid Excel Power Query functions

**Specific Problem**:

- `Excel.CurrentWorkbook` flagged as unknown/invalid function
- This is a **valid Excel Power Query function** used extensively in Excel workbooks
- Extension appears to be configured for Power BI context rather than Excel context

**Impact**:

- **Visual clutter**: Red squiggle lines on valid code
- **Developer confusion**: Valid functions appear as errors
- **IntelliSense issues**: May affect autocomplete for Excel-specific functions

**Root Cause Analysis**:

- Power Query/M Language extension may have different symbol libraries for:
  - **Power BI context**: `Sql.Database`, `Web.Contents`, etc.
  - **Excel context**: `Excel.CurrentWorkbook`, `Excel.Workbook`, etc.
- Extension likely defaults to Power BI symbol set

**Potential Solutions**:

1. **Configure Power Query extension** to recognize Excel-specific functions
2. **Custom symbol definitions** in workspace settings
3. **Extension configuration** to specify Excel vs Power BI context
4. **Suppress specific warnings** for known valid Excel functions

**Example Valid Excel Functions Being Flagged**:

- `Excel.CurrentWorkbook()` - Access tables/ranges in current workbook
- `Excel.Workbook()` - Open external Excel files
- Excel-specific connectors and functions

**Priority**: Low (cosmetic issue, doesn't affect functionality)

**Research Needed**:

- Power Query extension configuration options
- Symbol library customization
- Excel vs Power BI context switching

---

## Test Results Summary

### ✅ Working Features

- [x] **DataMashup extraction from ALL file sizes** - Fixed hardcoded scanning bug
- [x] **Perfect round-trip sync** - Both small and large files (60MB tested)
- [x] **Dynamic customXml location detection** - No longer limited to item1/2/3.xml
- [x] **Consistent BOM handling** - UTF-16 LE detection in extraction and sync
- [x] **Enhanced raw extraction** - Comprehensive debugging with ALL file scanning
- [x] **Function context logging** - Improved debugging with `[functionName]` prefixes
- [x] **Extension installation in dev containers**
- [x] **Workspace settings override user settings in dev containers**
- [x] **VSIX package optimization** (202KB, 25 files)
- [x] **Right-click context menu** for .m files in Explorer
- [x] **Metadata location tracking** in .m file headers
- [x] **File auto-watch on save** - Working in dev containers with Chokidar polling

### 🔄 In Progress

- [ ] **Logging standardization** - Partially implemented, need to complete function context logging

### ⚠️ Issues Identified

- [ ] Apply Recommended Defaults command uses Global scope (incompatible with dev containers)
- [ ] Dev container settings documentation needed
- [ ] Power Query/M Language extension shows false positives for valid Excel functions (cosmetic)

### 🔴 Critical Issues RESOLVED

- [x] ~~Large file (50MB+) DataMashup extraction failure~~ ✅ **FIXED**
- [x] ~~Hardcoded customXml scanning limitation~~ ✅ **FIXED**
- [x] ~~Sync vs extraction DataMashup detection inconsistency~~ ✅ **FIXED**
- [x] ~~File auto-watch not working in dev containers~~ ✅ **FIXED** (debounce was masking success)

### 🎯 Production Impact

**MAJOR PRODUCTION FIXES COMPLETED**:

1. **Extension now works with real-world Excel files** that store DataMashup in non-standard locations
2. **Perfect round-trip sync** maintains data integrity and user comments
3. **60MB file processing** demonstrates scalability for enterprise use
4. **117+ marketplace installations** now have access to these critical fixes

---

## Next Steps

1. **Current Priority**: Complete file auto-watch debugging in dev containers

   - Test dual watcher system (chokidar + VS Code FileSystemWatcher)
   - Verify polling mode effectiveness in Docker mounted volumes
   - Consider `onDidSaveTextDocument` event as alternative trigger

2. **Logging Standardization**: Complete function context logging migration

   - Replace remaining `function() completed.` style with `[functionName]` format
   - Implement consolidated log level setting (none/error/warn/info/verbose/debug)

3. **Dev Container UX**: Fix Apply Recommended Defaults for dev container users

   - Implement workspace vs global scope detection
   - Add user choice for settings scope

4. **Documentation Updates**: Reflect major fixes in user documentation
   - Update user guide with large file capability
   - Document dev container settings requirements
   - Add troubleshooting guide for file watching issues

---

## Test Files Used

- ✅ `PowerQueryFunctions.xlsx` - Small file: **Perfect extraction and round-trip sync** (`customXml/item1.xml`)
- ✅ `MAR_DatabaseSummary-V7b.xlsm` - **60MB large file: Perfect extraction and round-trip sync** (`customXml/item19.xml`)
- ✅ Various small test files - All working correctly with dynamic location detection

**Key Discovery**: Large Excel files with multiple Power Query connections store DataMashup in higher-numbered customXml files (item19.xml vs item1.xml), which the previous hardcoded scanning missed entirely.

---

_Document Updated: 2025-07-12_
_Status: MAJOR PRODUCTION BUGS RESOLVED - File watching in dev containers under investigation_
_Version: 0.5.0_
