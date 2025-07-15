# Debug Extraction Expected Results

This directory contains the expected results for debug extraction tests. Each subdirectory represents the expected output for a specific test file.

## Structure

```
debug-extraction/
├── simple/               # Expected results for simple.xlsx
│   ├── EXTRACTION_REPORT.json
│   └── item1_PowerQuery.m
├── complex/              # Expected results for complex.xlsm
│   ├── EXTRACTION_REPORT.json
│   └── item1_PowerQuery.m
├── binary/               # Expected results for binary.xlsb
│   ├── EXTRACTION_REPORT.json
│   └── item1_PowerQuery.m
└── no-powerquery/        # Expected results for no-powerquery.xlsx
    └── EXTRACTION_REPORT.json
```

## Expected Files

### Files with Power Query Content
For files containing Power Query definitions (`simple.xlsx`, `complex.xlsm`, `binary.xlsb`):

- **EXTRACTION_REPORT.json**: Comprehensive analysis report including:
  - File metadata (name, size, file count)
  - Scan summary (XML files scanned, DataMashup files found)
  - File breakdown by category
  - DataMashup source details
  - File categorization
  - Validation results
  - Recommendations

- **item1_PowerQuery.m**: Extracted M code from the DataMashup
  - Contains the actual Power Query M language code
  - Includes proper section and query definitions
  - May contain multiple shared expressions

- **DATAMASHUP_customXml_item1.xml**: Raw DataMashup XML (when generated)
- **customXml files**: Additional extracted XML files as needed

### Files without Power Query Content
For files with no Power Query content (`no-powerquery.xlsx`):

- **EXTRACTION_REPORT.json**: Analysis report showing:
  - File metadata
  - Scan results (0 DataMashup files found)
  - Empty datamashup_sources array
  - no_powerquery_content flag set to true
  - Appropriate recommendations

## Test Validation

Tests should validate:

1. **Report Structure**: EXTRACTION_REPORT.json contains all required fields
2. **M Code Content**: Power Query M files contain valid M language syntax
3. **File Counts**: Expected number of files generated
4. **Categorization**: Proper file type categorization
5. **Recommendations**: Appropriate recommendations for each file type

## Usage in Tests

The integration tests copy input files from `fixtures/` to `temp/`, run debug extraction, then compare the actual results against these expected results.

Expected results should be treated as read-only reference data and should not be modified during testing.
