# CSV Specific Scenarios

## Import Scenarios

- Import all columns from a standard comma-delimited file with headers.
- Import selected columns by Excel-style letters, such as `A`, `B`, `E`, `AA`.
- Import selected columns by header names, such as `ID`, `Name`, `Amount`.
- Exclude columns by letters.
- Exclude columns by header names.
- Import files without headers and generate column names like `A`, `B`, `C`.
- Import pipe-delimited files.
- Import tab-delimited files.
- Import semicolon-delimited files for European-style exports.
- Import custom multi-character delimiters.
- Import quoted fields that contain delimiters.
- Import quoted fields that contain escaped quotes.
- Import rows with missing trailing values.
- Import rows with extra trailing values.
- Trim fields.
- Preserve whitespace exactly.
- Choose encoding, such as UTF-8, UTF-8 BOM, Windows-1252, or UTF-16.
- Detect encoding from BOM.
- Skip empty lines.
- Keep empty lines as empty rows where needed.
- Return dictionaries.
- Return `DataTable`.
- Return raw row objects with source label and row number.
- Later: typed object mapping.
- Later: per-column type conversion and validation.
- Later: malformed row policy, such as throw, collect, or skip.

## Export Scenarios

- Export all columns to CSV.
- Export selected columns.
- Rename headers.
- Export without headers.
- Choose delimiter.
- Choose quote and escape behavior.
- Choose encoding.
- Format decimals, dates, booleans, and nulls.
- Stream large CSV output.
- Export `DataTable`.
- Export `IEnumerable<T>`.
- Export dictionaries.
- Export the same source as both `.xlsx` and `.csv`.

## Fixture Files

Current test fixtures:

- `mock-invoices-100.csv`
- `mock-invoices-1000.csv`
- `mock-invoices-10000.csv`
- `mock-invoices-50000.csv`
- `mock-wide-1000x75.csv`
- `mock-pipe-1000.csv`
- `mock-no-header-1000.csv`

These fixtures are designed to test:

- row count scaling
- wide column selection
- sparse letter selection
- header selection
- quoted delimiters
- quoted quotes
- custom delimiters
- no-header imports
