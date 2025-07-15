# Debug Extraction Test Results

This directory contains expected results for debug extraction tests.

## Structure

### simple.xlsx Debug Extraction
- **Expected files**: 5 files total
  - `EXTRACTION_REPORT.json` - Main analysis report
  - `item1_PowerQuery.m` - Extracted M code from DataMashup
  - `DATAMASHUP_customXml_item1.xml` - Raw DataMashup XML content
  - `customXml_item1.xml.txt` - Decoded customXml content
  - `customXml_itemProps1.xml.txt` - Decoded itemProps content

### complex.xlsm Debug Extraction  
- **Expected files**: 5 files total (same structure as simple.xlsx)
- **DataMashup location**: Should be found in `customXml/item1.xml`
- **M code**: Should contain valid Power Query M code with section declaration

### no-powerquery.xlsx Debug Extraction
- **Expected files**: 1 file total
  - `EXTRACTION_REPORT.json` - Analysis report showing no DataMashup found
- **No M code files**: Should not extract any .m files
- **Recommendations**: Should include "No DataMashup content found"

## Validation Criteria

### For files WITH Power Query:
1. `EXTRACTION_REPORT.json` must exist with valid structure
2. `dataMashupAnalysis.dataMashupFilesFound` > 0
3. At least one `*_PowerQuery.m` file must exist
4. M code must contain `section ` declaration
5. M code must be > 50 characters
6. Recommendations should include "Found DataMashup" and "Successfully extracted M code"

### For files WITHOUT Power Query:
1. `EXTRACTION_REPORT.json` must exist
2. `dataMashupAnalysis.dataMashupFilesFound` === 0  
3. No `*_PowerQuery.m` files should exist
4. Recommendations should include "No DataMashup content found"

## Test Implementation

Tests validate:
- File structure creation
- Report JSON structure and content
- M code extraction and validation
- Proper handling of files without Power Query
- Error conditions and edge cases
