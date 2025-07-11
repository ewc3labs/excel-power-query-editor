# Test Fixtures

This directory contains sample Excel files for testing the Excel Power Query Editor extension.

## Files

- `sample-with-powerquery.xlsx` - Sample Excel file containing Power Query for testing extraction and sync operations
- `sample-without-powerquery.xlsx` - Sample Excel file without Power Query for testing edge cases
- `test-data.csv` - Sample CSV data that can be referenced in Power Query scripts

## Usage

These files are used by the automated test suite to validate:
- Power Query extraction from Excel files
- Sync operations back to Excel
- Watch mode functionality
- Backup and cleanup operations

## Creating Test Files

To create new test Excel files with Power Query:
1. Open Excel
2. Go to Data > Get Data > From Other Sources > Blank Query
3. Create a simple M query like:
   ```m
   let
       Source = "Hello, World!",
       Result = Source
   in
       Result
   ```
4. Save the file in this directory
5. Update test cases to reference the new file
