# Volume Testing

VolumeRoundTripTests covers CSV and XLSX at 10k, 50k, 100k, 150k, 200k,
250k, 300k, 400k and 500k records, plus headers and one totals row.
Two 100-row smoke cases run normally. Large cases are skipped unless enabled.

Run from the repository root in 64-bit PowerShell:

    $env:DOCWEAVE_RUN_VOLUME_MATRIX = "1"
    dotnet test DocWeave.Tests/DocWeave.Tests.csproj -c Release -f net10.0 --filter "Category=VolumeMatrix"

For only 10k use filter "Category=VolumeMatrix&DisplayName~10000,".
Use DocWeaveNS.Tests/DocWeaveNS.Tests.csproj for the compatibility package.
Run packages/frameworks sequentially to avoid overlapping large workbooks.

The 20-column deterministic random dataset contains string IDs, dates, decimals,
doubles, integers, longs, booleans, quoted strings and multiline text. A fixed
seed makes runs repeatable. XLSX has date/monetary formatting and a styled totals
row with SUM formulas. CSV appends precomputed numeric totals; CSV cannot store
spreadsheet formatting. Totals are additional to the requested data-row count.

Export timing includes adaptation, rendering and file writing; CSV also includes
constructing totals. Import timing includes streaming and correctness assertions,
so this measures verified throughput rather than pure parser speed. Generation is
measured separately. Imports check 20 fields per record, all IDs and quoted
descriptions, nonempty dates, cost/revenue sums, row counts and totals markers.
CSV total amounts are checked. XLSX formulas are not evaluated by DocWeave.
Typed conversion of every date/double field is not asserted by this suite.

One JSONL file per case is written to volume-results under the test output folder.
It records UTC time, package, runtime, status, bytes and millisecond timings.
DOCWEAVE_VOLUME_RESULTS can specify an absolute results directory.
DOCWEAVE_KEEP_VOLUME_FILES=1 retains generated files; otherwise successful files
are deleted while timings remain. Failed files remain for diagnosis. A timing
phase is present only when it completed. Peak working set is the process lifetime
high-water mark rather than per-case allocation; isolate cases for comparisons.

The XLSX exporter materializes cells. The 500k x 20 case exceeds 10 million cells
and may require substantial RAM. These capacity tests have no arbitrary speed
threshold. Use Release builds and compare on consistent hardware/runtime/storage.
A smoke or 10k run does not imply the full matrix has been executed.