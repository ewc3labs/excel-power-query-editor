# Test Fixtures Documentation

This directory contains test Excel files for comprehensive end-to-end testing of the Excel Power Query Editor extension.

## Test Files

### Core Test Files
- **simple.xlsx** - Basic Power Query with single table import
- **complex.xlsm** - Multiple queries with dependencies and macros
- **binary.xlsb** - Binary format with Power Query content
- **no-powerquery.xlsx** - Excel file without any Power Query (edge case)

### Expected Outputs
The `expected/` directory contains the expected `.m` file content that should be extracted from each test file.

## Test Scenarios Covered

1. **Format Support**: .xlsx, .xlsm, .xlsb files
2. **Content Variety**: Simple vs complex Power Query scenarios
3. **Edge Cases**: Files without Power Query content
4. **Binary Format**: Specific testing for .xlsb handling

## Usage in Tests

Tests use these files to verify:
- Extraction produces expected .m content
- Sync operations work correctly
- Watch functionality operates properly
- Backup and cleanup functions work as expected
- Error handling for edge cases

Each test file should contain realistic Power Query scenarios that mirror real-world usage.
