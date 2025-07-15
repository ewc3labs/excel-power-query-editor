# Enhanced Debug Extraction Test Architecture

## Overview

This document describes the comprehensive test architecture implemented for validating the enhanced debug extraction functionality in the Excel Power Query Editor extension.

## Test Structure

### Fixtures Directory (`test/fixtures/`)
Contains read-only input files for testing:
- `simple.xlsx` - Basic Excel file with simple Power Query
- `complex.xlsm` - Complex Excel macro file with advanced Power Query
- `binary.xlsb` - Binary Excel file with Power Query
- `no-powerquery.xlsx` - Excel file with no Power Query content

### Expected Results (`test/fixtures/expected/debug-extraction/`)
Contains reference outputs for comparison:
```
debug-extraction/
├── simple/
│   ├── EXTRACTION_REPORT.json
│   └── item1_PowerQuery.m
├── complex/
│   ├── EXTRACTION_REPORT.json  
│   └── item1_PowerQuery.m
├── binary/
│   ├── EXTRACTION_REPORT.json
│   └── item1_PowerQuery.m
├── no-powerquery/
│   └── EXTRACTION_REPORT.json
└── README.md
```

### Temp Directory (`test/temp/`)
Working directory for test execution:
- Input files are copied here from fixtures
- Debug extraction operates on temp files
- Outputs are generated here and compared against expected results
- Cleaned up after each test run

## Test Methodology

### 1. Isolation Principle
- Tests copy input files from `fixtures/` to `temp/` before testing
- Fixtures directory remains clean and read-only
- No pollution of reference data with test outputs

### 2. Comprehensive Validation
Each test validates:
- ✅ Debug directory creation
- ✅ Required file generation (EXTRACTION_REPORT.json, M code files)
- ✅ Report structure and content
- ✅ M code syntax and structure
- ✅ File categorization accuracy
- ✅ Recommendation quality
- ✅ Comparison with expected results

### 3. Test Cases

#### Files with Power Query Content
**Test Files**: `simple.xlsx`, `complex.xlsm`, `binary.xlsb`

**Validation**:
- EXTRACTION_REPORT.json structure matches expected
- M code files generated with valid Power Query syntax
- DataMashup file count matches expected
- File categorization is accurate
- Recommendations are appropriate

#### Files without Power Query Content  
**Test Files**: `no-powerquery.xlsx`

**Validation**:
- EXTRACTION_REPORT.json generated with zero DataMashup files
- No M code files generated
- Appropriate "no Power Query" recommendations
- `no_powerquery_content` flag set correctly

## Integration Test Suite

Located in `test/integration.test.ts`:

### Enhanced Debug Extraction Tests Suite
```typescript
suite('Enhanced Debug Extraction Tests', () => {
    // Tests for files with Power Query content
    // Tests for files without Power Query content
    // Comprehensive validation and comparison
});
```

### Key Features
- **Timeout Management**: 10-second timeout for file processing
- **Error Handling**: Graceful handling of missing files
- **Detailed Logging**: Comprehensive console output for debugging
- **Expected Comparison**: Validation against reference results
- **Structure Validation**: Deep validation of report structure

## Validation Criteria

### EXTRACTION_REPORT.json
- File metadata (name, size, file count)
- Scan summary (XML files scanned, DataMashup files found)
- File breakdown by category
- DataMashup source details
- File categorization counts
- Validation results
- Recommendations array

### M Code Files
- Valid Power Query M syntax
- Section declarations present
- Appropriate file size (>50 characters for valid files)
- Correct file naming convention

### Directory Structure
- Debug directory created with correct naming
- All expected files generated
- No unexpected files or pollution

## Benefits

1. **Reliability**: Consistent validation across all test scenarios
2. **Maintainability**: Clear separation between inputs, outputs, and expectations
3. **Comprehensiveness**: Deep validation of all extraction aspects
4. **Non-Destructive**: Fixtures remain pristine for reproducible testing
5. **Debuggability**: Rich logging and clear error messages

## Usage

Run enhanced debug extraction tests:
```bash
npm test -- --grep "Enhanced Debug Extraction Tests"
```

Run all integration tests:
```bash
npm test
```

## Continuous Validation

The test suite ensures that:
- Debug extraction functionality remains stable across changes
- Report format consistency is maintained
- M code extraction quality is preserved
- Error handling for edge cases works correctly
- Performance characteristics remain acceptable

This architecture provides confidence in the debug extraction feature and enables safe refactoring and enhancement of the extraction logic.
