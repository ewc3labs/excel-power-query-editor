# Excel Power Query Editor - User Guide

## Overview
This VS Code extension allows you to extract Power Query formulas from Excel files, edit them as .m files, and sync changes back to Excel.

## Quick Start

### 1. Install the Extension
```bash
code --install-extension excel-power-query-editor-1.0.0.vsix
```

### 2. Extract Power Query from Excel
1. Right-click on any `.xlsx`, `.xlsm`, or `.xlsb` file in VS Code Explorer
2. Select "Extract Power Query from Excel"
3. Extension creates `filename_PowerQuery.m` in the same directory

### 3. Edit the .m File
- The generated .m file contains your Power Query formula
- Edit using VS Code's text editor (M syntax highlighting may be limited)
- Save changes as needed

### 4. Sync Changes Back to Excel
1. Right-click on the `.m` file
2. Select "Sync Power Query to Excel"
3. Extension automatically finds the corresponding Excel file and updates it

### 5. Automatic Sync (Optional)
1. Right-click on the `.m` file
2. Select "Watch File for Changes"
3. Extension automatically syncs any saved changes to Excel
4. Status bar shows watching indicator

## File Naming Convention

The extension uses a specific naming pattern:

- **Excel file**: `MyWorkbook.xlsx`
- **Power Query file**: `MyWorkbook_PowerQuery.m`

The sync feature automatically detects the Excel file by:
1. Removing `_PowerQuery` from the .m filename
2. Looking for Excel files with extensions `.xlsx`, `.xlsm`, or `.xlsb`
3. Checking the same directory first, then parent directories

## Supported Operations

### Context Menu Commands
- **Extract Power Query from Excel** - Extract .m files from Excel
- **Sync Power Query to Excel** - Update Excel with .m file changes
- **Watch File for Changes** - Enable automatic sync
- **Sync & Delete** - Sync to Excel and delete the .m file

### Command Palette
- `Excel PQ: Extract from Excel`
- `Excel PQ: Sync to Excel`
- `Excel PQ: Watch File`
- `Excel PQ: Stop Watching`
- `Excel PQ: Raw Extraction` (for debugging)

## File Structure Examples

### Simple Case
```
project/
├── MyWorkbook.xlsx
└── MyWorkbook_PowerQuery.m
```

### Subfolder Organization
```
project/
├── MyWorkbook.xlsx
└── MyWorkbook_PowerQuery/
    └── Section1.m
```

## Troubleshooting

### Sync Prompts for File Selection
If the sync command asks you to select an Excel file instead of auto-detecting:
1. Check that the Excel file exists in the expected location
2. Verify the naming convention matches (underscore, not dots)
3. Ensure the Excel file has a supported extension (`.xlsx`, `.xlsm`, `.xlsb`)

### No Power Query Found
If extraction reports no Power Query content:
- File may not contain Power Query formulas
- Try "Raw Extraction" to see all Excel content
- Check if the Excel file uses external data connections instead

### Watch Feature Not Working
- Check the status bar for watch indicators
- Ensure the Excel file path hasn't changed
- Stop watching and restart if needed

## File Extensions Supported

### Excel Files (Source)
- `.xlsx` - Excel Workbook
- `.xlsm` - Excel Macro-Enabled Workbook  
- `.xlsb` - Excel Binary Workbook

### Power Query Files (Generated)
- `.m` - Power Query M Language files

## Advanced Features

### Raw Extraction
For debugging or advanced analysis:
1. Right-click Excel file
2. Select "Raw Excel Extraction"
3. Creates `raw_excel_extraction/` folder with all Excel components

### Batch Operations
- The extension handles multiple Power Query definitions in a single Excel file
- Each query can be extracted to a separate .m file
- Sync operations can update specific queries

## Limitations

### Current Version Constraints
- This extension is compiled-only (no source code available)
- Limited to features already implemented
- Cannot be modified or extended

### Excel Integration
- Requires Excel files to be closed during sync operations
- Changes are written to Excel files directly
- Always creates backup files (.backup.timestamp)

## Tips for Best Results

1. **Close Excel before syncing** - Avoid file locking issues
2. **Use descriptive query names** - Makes .m files easier to identify
3. **Regular backups** - Extension creates backups, but additional backups recommended
4. **Test sync on copies** - Verify results before working on important files
5. **Watch status bar** - Provides feedback on watch and sync operations

## Credits

This extension integrates with:
- [excel-datamashup](https://github.com/Vladinator/excel-datamashup) library by Vladinator (GPL-3.0)
- Recommended: [GrapeCity Excel Viewer](https://marketplace.visualstudio.com/items?itemName=grapecity.gc-excelviewer) for Excel file preview

---

*For technical issues or questions, refer to the project's README.md file.*
