using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Pivot tables: row, column, filter and value fields, grand totals, several pivots per source, and the layouts that are rejected.
/// </summary>
public sealed class ExcelPivotExportTests
{
    [Fact]
    public void _20_Export_pivot_table_from_source_data_on_same_sheet()
    {
        var table = CreatePivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Sales Pivot", sheet =>
            {
                sheet.AddCell("A1", "Sales source with same-sheet pivot", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(table)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#D9E2F3"));
                sheet.AddPivotTable("SalesByRegion", pivot => pivot
                    .Source("Sales Pivot", "$A$3:$E$11")
                    .StartAt("H3")
                    .Rows("Region")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .Style("PivotStyleMedium9"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_20_Export_pivot_table_from_source_data_on_same_sheet), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Sales Pivot"));
        Assert.Equal(1, CountPivotCacheDefinitionParts(bytes));
    }

    [Fact]
    public void _21_Export_pivot_table_from_source_data_on_different_sheet()
    {
        var table = CreatePivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Sales Data", sheet => sheet
                .AddCell("A1", "Sales source data", style => style.Bold().FontSize(16).FontColor("#385723"))
                .AddTable(table)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#385723").Border())
                .AlternateRowStyle(style => style.BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#A9D18E")))
            .AddSheet("Pivot Summary", sheet => sheet
                .AddCell("A1", "Pivot summary from another worksheet", style => style.Bold().FontSize(16).FontColor("#7030A0"))
                .AddPivotTable("SalesByRegionAndProduct", pivot => pivot
                    .Source("Sales Data", "$A$3:$E$11")
                    .StartAt("A3")
                    .Rows("Region")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .ShowGrandTotals(rows: true, columns: false)
                    .Style("PivotStyleLight16")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_21_Export_pivot_table_from_source_data_on_different_sheet), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Pivot Summary"));
        Assert.Equal(1, CountPivotCacheDefinitionParts(bytes));
    }

    [Fact]
    public void _22_Export_pivot_table_with_multiple_row_and_column_fields()
    {
        var table = CreatePivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Sales Data", sheet => sheet
                .AddTable(table)
                .StartAt("A1"))
            .AddSheet("Pivot Summary", sheet => sheet
                .AddPivotTable("SalesByRegionProductQuarterChannel", pivot => pivot
                    .Source("Sales Data", "$A$1:$E$9")
                    .StartAt("A1")
                    .Rows("Region", "Product")
                    .Columns("Quarter", "Channel")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .Style("PivotStyleMedium9")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_22_Export_pivot_table_with_multiple_row_and_column_fields), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Pivot Summary"));
        Assert.Equal(1, CountPivotCacheDefinitionParts(bytes));

        var definition = GetPivotTableDefinition(bytes, "Pivot Summary");

        Assert.Equal(2, definition.RowFields!.Elements<Field>().Count());
        Assert.Equal(2, definition.ColumnFields!.Elements<Field>().Count());

        // Two column fields need two header rows stacked above the data. Reserving only one,
        // as the single-column-field formula used to, would leave the pivot table one row short.
        var (start, end) = ParseReferenceRange(definition.Location!.Reference!.Value!);
        Assert.Equal(11, end.row - start.row + 1);
        Assert.Equal(7, end.column - start.column + 1);
    }

    [Fact]
    public void _23_Export_dashboard_with_multiple_pivot_tables_from_one_source()
    {
        var table = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Detailed Sales", sheet => sheet
                .AddCell("A1", "Detailed sales source", style => style.Bold().FontSize(16).FontColor("#1F4E78"))
                .AddTable(table)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .AlternateRowStyle(style => style.BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Thin, "#9EADCC")))
            .AddSheet("Pivot Dashboard", sheet =>
            {
                sheet.AddCell("A1", "Sales Pivot Dashboard", style => style.Bold().FontSize(16).FontColor("#7030A0"));
                sheet.AddPivotTable("RegionProductByQuarter", pivot => pivot
                    .Source("Detailed Sales", "$A$3:$I$15")
                    .StartAt("A3")
                    .Rows("Region", "Product")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .Style("PivotStyleMedium4"));
                sheet.AddPivotTable("ChannelSegmentByQuarter", pivot => pivot
                    .Source("Detailed Sales", "$A$3:$I$15")
                    .StartAt("H3")
                    .Rows("Channel", "Segment")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .ShowGrandTotals(rows: true, columns: false)
                    .Style("PivotStyleLight16"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_23_Export_dashboard_with_multiple_pivot_tables_from_one_source), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(2, CountPivotTableParts(bytes, "Pivot Dashboard"));
        Assert.Equal(2, CountPivotCacheDefinitionParts(bytes));

        var definitions = GetPivotTableDefinitions(bytes, "Pivot Dashboard");

        Assert.Contains(definitions, definition => definition.Name == "RegionProductByQuarter");
        Assert.Contains(definitions, definition => definition.Name == "ChannelSegmentByQuarter");
        Assert.All(definitions, definition => Assert.Equal(1, definition.ColumnFields!.Elements<Field>().Count()));
    }

    [Fact]
    public void _24_Export_multi_sheet_pivot_pack_from_one_source()
    {
        var table = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Detailed Sales", sheet => sheet
                .AddCell("A1", "Detailed sales source", style => style.Bold().FontSize(16).FontColor("#385723"))
                .AddTable(table)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#385723").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#A9D18E")))
            .AddSheet("Regional Pivot", sheet => sheet
                .AddCell("A1", "Region and product sales by quarter", style => style.Bold().FontSize(16).FontColor("#1F4E78"))
                .AddPivotTable("RegionalProductSales", pivot => pivot
                    .Source("Detailed Sales", "$A$3:$I$15")
                    .StartAt("A3")
                    .Rows("Region", "Product")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .ShowGrandTotals(rows: true, columns: true)
                    .Style("PivotStyleMedium9")))
            .AddSheet("Channel Pivot", sheet => sheet
                .AddCell("A1", "Channel and segment sales by quarter", style => style.Bold().FontSize(16).FontColor("#7030A0"))
                .AddPivotTable("ChannelSegmentSales", pivot => pivot
                    .Source("Detailed Sales", "$A$3:$I$15")
                    .StartAt("A3")
                    .Rows("Channel", "Segment")
                    .Columns("Quarter")
                    .Values(values => values.Sum("Amount", "Total Sales"))
                    .ShowGrandTotals(rows: true, columns: false)
                    .Style("PivotStyleLight16")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_24_Export_multi_sheet_pivot_pack_from_one_source), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Regional Pivot"));
        Assert.Equal(1, CountPivotTableParts(bytes, "Channel Pivot"));
        Assert.Equal(2, CountPivotCacheDefinitionParts(bytes));

        var regionalDefinition = GetPivotTableDefinition(bytes, "Regional Pivot");
        var channelDefinition = GetPivotTableDefinition(bytes, "Channel Pivot");

        Assert.Equal(2, regionalDefinition.RowFields!.Elements<Field>().Count());
        Assert.Equal(1, regionalDefinition.ColumnFields!.Elements<Field>().Count());
        Assert.Equal(2, channelDefinition.RowFields!.Elements<Field>().Count());
        Assert.Equal(1, channelDefinition.ColumnFields!.Elements<Field>().Count());

        var (regionalStart, regionalEnd) = ParseReferenceRange(regionalDefinition.Location!.Reference!.Value!);
        var (channelStart, channelEnd) = ParseReferenceRange(channelDefinition.Location!.Reference!.Value!);

        Assert.True(regionalEnd.row > regionalStart.row);
        Assert.True(regionalEnd.column > regionalStart.column);
        Assert.True(channelEnd.row > channelStart.row);
        Assert.True(channelEnd.column > channelStart.column);
    }

    [Fact]
    public void _46_Export_pivot_table_with_filters_and_multiple_value_fields()
    {
        var source = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet
                .AddTable(source)
                .StartAt("A1")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7")))
            .AddSheet("Pivot", sheet => sheet
                .AddPivotTable("FilteredMultiValuePivot", pivot => pivot
                    .Source("Source", "A1:I13")
                    .StartAt("B3")
                    .Filters("Channel")
                    .Rows("Region", "Product")
                    .Columns("Quarter")
                    .Values(values => values
                        .Sum("Amount", "Sales")
                        .Average("Cost", "Average Cost"))
                    .Style("PivotStyleMedium4")
                    .ShowGrandTotals(rows: true, columns: true)))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_46_Export_pivot_table_with_filters_and_multiple_value_fields), bytes);

        var definition = GetPivotTableDefinition(bytes, "Pivot");

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Pivot"));
        Assert.Equal(1U, definition.PageFields!.Count!.Value);
        Assert.Equal(2U, definition.DataFields!.Count!.Value);
        Assert.Equal(
            ["Sales", "Average Cost"],
            definition.DataFields.Elements<DataField>().Select(field => field.Name!.Value));
    }

    [Fact]
    public void _48_Export_grouped_pivot_table_with_expandable_region_product_rows()
    {
        var source = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet
                .AddTable(source)
                .StartAt("A1")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7")))
            .AddSheet("Grouped Pivot", sheet => sheet
                .AddPivotTable("GroupedRegionProductPivot", pivot => pivot
                    .Source("Source", "A1:I13")
                    .StartAt("B3")
                    .Filters("Channel")
                    .Rows("Region", "Product")
                    .Columns("Quarter")
                    .Values(values => values
                        .Sum("Amount", "Sales")
                        .Average("Cost", "Average Cost"))
                    .Style("PivotStyleMedium4")
                    .ShowGrandTotals(rows: true, columns: true)))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_48_Export_grouped_pivot_table_with_expandable_region_product_rows), bytes);

        var definition = GetPivotTableDefinition(bytes, "Grouped Pivot");

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Grouped Pivot"));

        // Source headers: Region 0, Country 1, Product 2, Segment 3, Quarter 4, Channel 5, Amount 6, Units 7, Cost 8.
        Assert.Equal([0, 2], definition.RowFields!.Elements<Field>().Select(field => field.Index!.Value));
        Assert.Equal([5], definition.PageFields!.Elements<PageField>().Select(field => field.Field!.Value));

        // Quarter plus the "Values" pseudo field (-2) that Excel adds when there is more than one value.
        Assert.Equal([4, -2], definition.ColumnFields!.Elements<Field>().Select(field => field.Index!.Value));
        Assert.Equal(2U, definition.ColumnFields.Count!.Value);

        // 3 quarters x 2 values, plus one grand total column per value.
        var columnItems = definition.ColumnItems!.Elements<RowItem>().ToList();
        Assert.Equal(8U, definition.ColumnItems.Count!.Value);
        Assert.Equal(8, columnItems.Count);
        Assert.Equal([0U, 1U, 0U, 1U, 0U, 1U, 0U, 1U], columnItems.Select(item => item.Index?.Value ?? 0U));
        Assert.Equal(2, columnItems.Count(item => item.ItemType?.Value == ItemValues.Grand));

        // The last "x" of each item names its value field, matching the item's "i" attribute.
        Assert.All(columnItems, item => Assert.Equal(item.Index?.Value ?? 0U, GetItemValues(item).Last()));

        // 11 Region/Product combinations plus the grand total row.
        var rowItems = definition.RowItems!.Elements<RowItem>().ToList();
        Assert.Equal(12, rowItems.Count);

        // Products stay grouped under their region: Europe (0) has 3 rows, Americas (1) 3, EMEA (2) 2, APAC (3) 3.
        var regionIndexes = rowItems.Where(item => item.ItemType?.Value != ItemValues.Grand).Select(item => GetItemValues(item)[0]).ToList();
        Assert.Equal([0U, 0U, 0U, 1U, 1U, 1U, 2U, 2U, 3U, 3U, 3U], regionIndexes);

        // Filter row above (row 1), a blank row, then the table: 3 header rows, 11 rows, grand total.
        Assert.Equal("B3:K17", definition.Location!.Reference!.Value);
        Assert.Equal(1U, definition.Location.RowPageCount!.Value);
        Assert.Equal(1U, definition.Location.FirstHeaderRow!.Value);
        Assert.Equal(3U, definition.Location.FirstDataRow!.Value);
    }

    [Fact]
    public void _49_Export_pivot_table_with_several_values_and_no_column_fields()
    {
        var source = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet.AddTable(source).StartAt("A1"))
            .AddSheet("Region Totals", sheet => sheet
                .AddPivotTable("RegionTotals", pivot => pivot
                    .Source("Source", "A1:I13")
                    .StartAt("A3")
                    .Rows("Region")
                    .Values(values => values
                        .Sum("Amount", "Sales")
                        .Sum("Units", "Units Sold")
                        .Max("Cost", "Peak Cost"))))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_49_Export_pivot_table_with_several_values_and_no_column_fields), bytes);

        var definition = GetPivotTableDefinition(bytes, "Region Totals");

        Xlsx.AssertValid(bytes);

        // Without column fields the value captions are the only columns: just the "Values" pseudo field.
        Assert.Equal([-2], definition.ColumnFields!.Elements<Field>().Select(field => field.Index!.Value));
        var columnItems = definition.ColumnItems!.Elements<RowItem>().ToList();
        Assert.Equal([0U, 1U, 2U], columnItems.Select(item => item.Index?.Value ?? 0U));
        Assert.DoesNotContain(columnItems, item => item.ItemType?.Value == ItemValues.Grand);
        Assert.Equal(3U, definition.DataFields!.Count!.Value);

        // The header is the caption row itself, so data starts on the next row: A3:D8 = header, 4 regions, grand total.
        Assert.Equal(0U, definition.Location!.FirstHeaderRow!.Value);
        Assert.Equal(1U, definition.Location.FirstDataRow!.Value);
        Assert.Equal("A3:D8", definition.Location.Reference!.Value);
    }

    [Fact]
    public void _50_Export_pivot_table_with_two_values_from_the_same_source_field()
    {
        var source = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet.AddTable(source).StartAt("A1"))
            .AddSheet("Pivot", sheet => sheet
                .AddPivotTable("SumAndAverage", pivot => pivot
                    .Source("Source", "A1:I13")
                    .StartAt("A3")
                    .Rows("Region")
                    .Columns("Quarter")
                    .Values(values => values
                        .Sum("Amount", "Total Sales")
                        .Average("Amount", "Average Sale"))))
            .ToBytes();

        var definition = GetPivotTableDefinition(bytes, "Pivot");
        var dataFields = definition.DataFields!.Elements<DataField>().ToList();

        Xlsx.AssertValid(bytes);
        Assert.Equal(2, dataFields.Count);
        Assert.All(dataFields, field => Assert.Equal(6U, field.Field!.Value));
        Assert.Equal(DataConsolidateFunctionValues.Sum, dataFields[0].Subtotal!.Value);
        Assert.Equal(DataConsolidateFunctionValues.Average, dataFields[1].Subtotal!.Value);

        // The shared source field is flagged as a data field exactly once.
        Assert.Equal(1, definition.PivotFields!.Elements<PivotField>().Count(field => field.DataField?.Value == true));
    }

    [Fact]
    public void _51_Export_pivot_table_without_grand_totals_omits_grand_total_items()
    {
        var source = CreateComplexPivotSalesTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet.AddTable(source).StartAt("A1"))
            .AddSheet("Pivot", sheet => sheet
                .AddPivotTable("NoTotals", pivot => pivot
                    .Source("Source", "A1:I13")
                    .StartAt("A3")
                    .Rows("Region")
                    .Columns("Quarter")
                    .Values(values => values
                        .Sum("Amount", "Sales")
                        .Sum("Cost", "Cost Total"))
                    .ShowGrandTotals(rows: false, columns: false)))
            .ToBytes();

        var definition = GetPivotTableDefinition(bytes, "Pivot");

        Xlsx.AssertValid(bytes);
        Assert.False(definition.RowGrandTotals!.Value);
        Assert.False(definition.ColumnGrandTotals!.Value);

        // 3 quarters x 2 values and no grand total columns; 4 regions and no grand total row.
        Assert.Equal(6, definition.ColumnItems!.Elements<RowItem>().Count());
        Assert.Equal(4, definition.RowItems!.Elements<RowItem>().Count());
        Assert.DoesNotContain(definition.ColumnItems.Elements<RowItem>(), item => item.ItemType?.Value == ItemValues.Grand);
        Assert.DoesNotContain(definition.RowItems.Elements<RowItem>(), item => item.ItemType?.Value == ItemValues.Grand);
    }

    [Fact]
    public void _52_Pivot_table_rejects_layouts_that_excel_would_repair()
    {
        var source = CreateComplexPivotSalesTable();

        byte[] Build(Action<ExcelPivotTableBuilder> configure) => ExcelWriter.Create()
            .AddSheet("Source", sheet => sheet.AddTable(source).StartAt("A1"))
            .AddSheet("Pivot", sheet => sheet.AddPivotTable("Broken", configure))
            .ToBytes();

        // A value caption may not repeat a source field name.
        var sameAsField = Assert.Throws<InvalidOperationException>(() => Build(pivot => pivot
            .Source("Source", "A1:I13").StartAt("A3").Rows("Region")
            .Values(values => values.Sum("Amount", "Amount"))));
        Assert.Contains("matches a source field name", sameAsField.Message);

        // Value captions must be unique.
        var duplicateCaption = Assert.Throws<InvalidOperationException>(() => Build(pivot => pivot
            .Source("Source", "A1:I13").StartAt("A3").Rows("Region")
            .Values(values => values.Sum("Amount", "Total").Sum("Cost", "total"))));
        Assert.Contains("more than one value field captioned", duplicateCaption.Message);

        // A field can only sit on one axis.
        var twoAxes = Assert.Throws<InvalidOperationException>(() => Build(pivot => pivot
            .Source("Source", "A1:I13").StartAt("A3").Rows("Region").Columns("Region")
            .Values(values => values.Sum("Amount", "Sales"))));
        Assert.Contains("both the row and column axes", twoAxes.Message);

        // Filters need room above the table: one row per filter plus a blank row.
        var noRoomForFilter = Assert.Throws<InvalidOperationException>(() => Build(pivot => pivot
            .Source("Source", "A1:I13").StartAt("A2").Filters("Channel").Rows("Region")
            .Values(values => values.Sum("Amount", "Sales"))));
        Assert.Contains("start at row 3", noRoomForFilter.Message);
    }
}
