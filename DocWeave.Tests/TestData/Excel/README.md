# Excel test fixtures

Source workbooks for the Excel import tests. They are inputs: the tests read them and never write them. (Workbooks that the export
tests generate go to `TestOutput/Excel`, which is git-ignored and rebuilt on every run.)

All of them were created with DocWeave's own exporter, except `stream-shared-string-table.xlsx`, which was built by hand with the
Open XML SDK so that its text lives in a shared string table, the way Excel stores it. Every sheet starts with a header row in
row 1 unless the description says otherwise. To change one, edit it in Excel and keep the layout described here, or the test that
reads it will need updating too.

## Typed objects (`ExcelObjectImportTests`)

| File | Test | Contents |
|---|---|---|
| `objects-convention-headers.xlsx` | `_53_` | Sheet `Invoices`. Title in B2, table from B4: headers `Invoice No`, `customer`, `net_amount`; 2 rows (INV-001 Alice 125.50, INV-002 Bob 300). |
| `objects-attribute-mapping.xlsx` | `_54_` | Headers `Invoice Number`, `Bought By`, `Net`; 1 row (INV-001, Alice, 125.50). |
| `objects-start-cell-b4.xlsx` | `_55_` | Column A holds a note in row 4. The table starts at B4: `Invoice Number`, `Whatever the header says`; 1 row (INV-001, Alice). |
| `objects-excel-date-numbers.xlsx` | `_57_` | Headers `When`, `Day`; 1 row of Excel date numbers: 45123.5 and 45123 (16 July 2023). |
| `objects-good-and-bad-rows.xlsx` | `_58_` | Headers `Count`, `Paid`, `Status`, `Date`; 4 text rows. Row 2 is fine, row 3 has `Paid` = maybe, row 4 has `Status` = Lost and `Date` = not a date, row 5 is fine. |
| `objects-blank-values.xlsx` | `_59_` | Headers `Count`, `OptionalCount`, `Ratio`; 1 row, `Count` and `OptionalCount` blank, `Ratio` 1. |
| `objects-whole-number-limits.xlsx` | `_60_` | Header `Count`; rows 3.5, 3000000000, 3. |
| `objects-missing-required-column.xlsx` | `_61_` | Header `Label`; rows one, two. |
| `objects-blank-required-value.xlsx` | `_62_` | Headers `Code`, `Label`; row 2 has Code A and a blank Label, row 3 has a blank Code and Label `second`. |
| `objects-unsupported-property-type.xlsx` | `_63_` | Header `Tags`; 1 row (a). |
| `objects-one-bad-count.xlsx` | `_64_` | Header `Count` (text); rows 1, x, 3. |
| `objects-progress-and-audit.xlsx` | `_65_` | Header `Count` (text); rows 1, secret-value. |
| `objects-header-only-sheet.xlsx` | `_66_` | Sheet `Empty` with one cell, A1 = `Something else`. |
| `records-invoice-columns.xlsx` | `_67_` | Headers `Invoice No`, `Net Amount`, `Date`; 1 row (INV-001, 12.5, Excel date number 45123). |
| `records-reps-blank-age.xlsx` | `_68_` | Headers `Sales Rep`, `Age`; rows (Ada, 36), (Grace, blank), (blank, 50). |
| `objects-constructor-validation.xlsx` | `_69_` | Header `Value`; rows 3, -1, 5. |
| `objects-fluent-map-invoices.xlsx` | `_70_` | Headers `Number`, `Buyer`, `Amount`; rows (INV-001, Alice, 125.50), (INV-002, blank, 300). |
| `objects-column-selection.xlsx` | `_71_` | Headers `Number`, `Internal`, `Buyer`; rows (INV-001, hidden, Alice), (INV-002, hidden, Bob). |

`_56_` (export typed objects and import them back) generates its own workbook, because the export is what it tests.

## Streaming rows (`ExcelStreamingImportTests`)

| File | Tests | Contents |
|---|---|---|
| `stream-invoices-25-rows.xlsx` | `_72_` | Sheet `Invoices`, headers `InvoiceNo`, `Customer`, `Net`; 25 rows. |
| `stream-invoices-with-preamble.xlsx` | `_73_` | Text in B2 and B3, then the table from B5 (3 rows: Customer 1 to 3). |
| `stream-cell-right-of-header.xlsx` | `_74_`, `_75_`, `_76_` | The 3-column invoice table with 2 rows, plus a value `reviewed` in D2, which has no header. |
| `stream-invoices-2-rows.xlsx` | `_77_` | The invoice table with 2 rows. |
| `stream-header-only.xlsx` | `_78_` | The invoice table's header row and no data. |
| `stream-invoices-50-rows.xlsx` | `_79_` | The invoice table with 50 rows. |
| `stream-invoices-100-rows.xlsx` | `_80_` | The invoice table with 100 rows. |
| `stream-shared-string-table.xlsx` | `_81_` | Column A only, rows 1 to 4, text stored in the shared string table: London, Europe, United States, Europe. |
| `stream-sparse-sheet.xlsx` | `_82_` | `Region` in A1 and `Income` in C1, column B empty; data in rows 2 and 6 only (London 1500, Asia 980). |
