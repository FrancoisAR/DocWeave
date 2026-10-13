using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Writes pivot cache definitions, cache records and pivot table definitions for the pivot tables on a worksheet.
/// </summary>
internal static class ExcelPivotRenderer
{
    private sealed record PivotSourceData(IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<object?>> Rows);

    private sealed record PivotAxisKey(string Key, IReadOnlyList<int> Indexes);

    internal static void AddPivotTables(
        WorkbookPart workbookPart,
        RenderedWorksheetContext renderedSheet,
        IReadOnlyDictionary<string, RenderedWorksheetContext> renderedSheets,
        ref uint nextPivotCacheId)
    {
        var sheetSpec = renderedSheet.SheetSpec;
        for (var i = 0; i < sheetSpec.PivotTables.Count; i++)
        {
            var pivotTable = sheetSpec.PivotTables[i];
            if (pivotTable.Source is null)
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' does not have a source range.");
            }

            if (pivotTable.ValueFields.Count == 0)
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' must have at least one value field.");
            }

            // The cache is a copy of the source data, taken from the cells already rendered for the source sheet.
            // That is why the source sheet has to be part of this workbook.
            var sourceData = ResolvePivotSourceData(pivotTable, renderedSheets);
            ValidatePivotDefinition(pivotTable, sourceData.Headers);

            var cacheId = nextPivotCacheId++;
            var cacheDefinitionPart = AddPivotCacheDefinition(workbookPart, pivotTable, sourceData, cacheId);
            var pivotTablePart = renderedSheet.WorksheetPart.AddNewPart<PivotTablePart>();
            // Link the pivot table to its cache definition, using a fixed relationship id.
            pivotTablePart.AddPart(cacheDefinitionPart, "rId1");
            pivotTablePart.PivotTableDefinition = CreatePivotTableDefinition(pivotTable, sourceData, cacheId);
            pivotTablePart.PivotTableDefinition.Save();
        }
    }

    private static PivotTableCacheDefinitionPart AddPivotCacheDefinition(
        WorkbookPart workbookPart,
        ExcelPivotTableSpec pivotTable,
        PivotSourceData sourceData,
        uint cacheId)
    {
        var cacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var recordsPart = cacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>();

        recordsPart.PivotCacheRecords = CreatePivotCacheRecords(sourceData);
        recordsPart.PivotCacheRecords.Save();

        cacheDefinitionPart.PivotCacheDefinition = new PivotCacheDefinition
        {
            Id = cacheDefinitionPart.GetIdOfPart(recordsPart),
            SaveData = true,
            RefreshOnLoad = pivotTable.RefreshOnOpen,
            EnableRefresh = true,
            CreatedVersion = 6,
            RefreshedVersion = 6,
            MinRefreshableVersion = 3,
            RecordCount = (uint)sourceData.Rows.Count
        };
        cacheDefinitionPart.PivotCacheDefinition.Append(new CacheSource
        {
            Type = SourceValues.Worksheet,
            WorksheetSource = new WorksheetSource
            {
                Sheet = ExcelNames.SanitizeSheetName(pivotTable.Source!.SheetName),
                Reference = pivotTable.Source.Range.Replace("$", string.Empty, StringComparison.Ordinal)
            }
        });
        cacheDefinitionPart.PivotCacheDefinition.Append(CreateCacheFields(sourceData));
        cacheDefinitionPart.PivotCacheDefinition.Save();

        // The workbook lists its pivot caches in one PivotCaches element, placed straight after Sheets. Create it on first use.
        var pivotCaches = workbookPart.Workbook.GetFirstChild<PivotCaches>();
        if (pivotCaches is null)
        {
            pivotCaches = new PivotCaches();
            workbookPart.Workbook.InsertAfter(pivotCaches, workbookPart.Workbook.Sheets);
        }

        pivotCaches.Append(new PivotCache
        {
            CacheId = cacheId,
            Id = workbookPart.GetIdOfPart(cacheDefinitionPart)
        });

        return cacheDefinitionPart;
    }

    private static CacheFields CreateCacheFields(PivotSourceData sourceData)
    {
        var cacheFields = new CacheFields { Count = (uint)sourceData.Headers.Count };
        for (var columnIndex = 0; columnIndex < sourceData.Headers.Count; columnIndex++)
        {
            var values = sourceData.Rows
                .Select(row => columnIndex < row.Count ? row[columnIndex] : null)
                .DistinctBy(FormatPivotCacheKey)
                .ToList();

            cacheFields.Append(new CacheField
            {
                Name = sourceData.Headers[columnIndex],
                SharedItems = CreateSharedItems(values)
            });
        }

        return cacheFields;
    }

    private static SharedItems CreateSharedItems(IReadOnlyList<object?> values)
    {
        // A column with numbers is flagged as numeric, and as integer only when every non-blank value is a whole number.
        var hasNumbers = values.Any(value => value is double or decimal or int or long);
        var sharedItems = new SharedItems { Count = (uint)values.Count };
        if (hasNumbers)
        {
            sharedItems.ContainsSemiMixedTypes = false;
            sharedItems.ContainsString = false;
            sharedItems.ContainsNumber = true;
            sharedItems.ContainsInteger = values.All(value => value is null || IsWholeNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        }

        foreach (var value in values)
        {
            sharedItems.Append(CreatePivotCacheValue(value));
        }

        return sharedItems;
    }

    private static bool IsWholeNumber(double value)
    {
        return Math.Abs(value % 1) < double.Epsilon;
    }

    private static PivotCacheRecords CreatePivotCacheRecords(PivotSourceData sourceData)
    {
        var sharedItemsByColumn = Enumerable
            .Range(0, sourceData.Headers.Count)
            .Select(columnIndex => sourceData.Rows
                .Select(row => columnIndex < row.Count ? row[columnIndex] : null)
                .DistinctBy(FormatPivotCacheKey)
                .ToList())
            .ToList();
        var records = new PivotCacheRecords { Count = (uint)sourceData.Rows.Count };
        foreach (var row in sourceData.Rows)
        {
            var record = new PivotCacheRecord();
            for (var columnIndex = 0; columnIndex < sourceData.Headers.Count; columnIndex++)
            {
                record.Append(CreatePivotCacheRecordValue(
                    columnIndex < row.Count ? row[columnIndex] : null,
                    sharedItemsByColumn[columnIndex]));
            }

            records.Append(record);
        }

        return records;
    }

    private static OpenXmlElement CreatePivotCacheRecordValue(object? value, IReadOnlyList<object?> sharedItems)
    {
        if (value is null)
        {
            return new MissingItem();
        }

        // Numbers are written into the record as they are. Any other value is written as an index into the column's shared items.
        if (value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal)
        {
            return new NumberItem { Val = Convert.ToDouble(value, CultureInfo.InvariantCulture) };
        }

        // The shared items are the column's distinct values, in first-seen order. IndexOfPivotValue finds the position in that list.
        return new FieldItem { Val = (uint)IndexOfPivotValue(sharedItems, value) };
    }

    private static OpenXmlElement CreatePivotCacheValue(object? value)
    {
        return value switch
        {
            null => new MissingItem(),
            bool boolValue => new BooleanItem { Val = boolValue },
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal => new NumberItem { Val = Convert.ToDouble(value, CultureInfo.InvariantCulture) },
            _ => new StringItem { Val = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty }
        };
    }

    private static PivotSourceData ResolvePivotSourceData(
        ExcelPivotTableSpec pivotTable,
        IReadOnlyDictionary<string, RenderedWorksheetContext> renderedSheets)
    {
        var source = pivotTable.Source!;
        var sourceSheetName = ExcelNames.SanitizeSheetName(source.SheetName);
        if (!renderedSheets.TryGetValue(sourceSheetName, out var sourceSheet))
        {
            throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' source sheet '{source.SheetName}' was not found.");
        }

        // The first row of the range is the header row. The rows below it are the data.
        var (start, end) = ExcelCellAddress.ParseRange(source.Range);
        var startRow = Math.Min(start.Row, end.Row);
        var endRow = Math.Max(start.Row, end.Row);
        var startColumn = Math.Min(start.Column, end.Column);
        var endColumn = Math.Max(start.Column, end.Column);

        var headers = new List<string>();
        for (var column = startColumn; column <= endColumn; column++)
        {
            headers.Add(Convert.ToString(GetRenderedCellValue(sourceSheet.Rows, startRow, column), CultureInfo.InvariantCulture) ?? string.Empty);
        }

        if (headers.Any(string.IsNullOrWhiteSpace))
        {
            throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' source range '{source.Range}' must include non-empty headers in the first row.");
        }

        var rows = new List<IReadOnlyList<object?>>();
        for (var row = startRow + 1; row <= endRow; row++)
        {
            var values = new List<object?>();
            for (var column = startColumn; column <= endColumn; column++)
            {
                values.Add(GetRenderedCellValue(sourceSheet.Rows, row, column));
            }

            rows.Add(values);
        }

        ValidatePivotFields(pivotTable, headers);
        return new PivotSourceData(headers, rows);
    }

    /// <summary>
    /// Rejects layouts Excel would repair or silently rewrite, so callers get a clear error instead of a damaged workbook.
    /// </summary>
    private static void ValidatePivotDefinition(ExcelPivotTableSpec pivotTable, IReadOnlyList<string> headers)
    {
        // Excel lets a source field sit on one axis only (rows, columns, or report filters).
        var placedFields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (axis, fields) in new[]
        {
            ("row", pivotTable.RowFields),
            ("column", pivotTable.ColumnFields),
            ("filter", pivotTable.FilterFields)
        })
        {
            foreach (var field in fields)
            {
                if (!placedFields.TryAdd(field, axis))
                {
                    throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' places field '{field}' on both the {placedFields[field]} and {axis} axes. A field can only be used once.");
                }
            }
        }

        // A value caption that repeats a source field name, or another caption, makes Excel repair the file.
        var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var valueField in pivotTable.ValueFields)
        {
            if (string.IsNullOrWhiteSpace(valueField.Caption))
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' has a value field for '{valueField.FieldName}' without a caption.");
            }

            if (headers.Contains(valueField.Caption, StringComparer.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' value caption '{valueField.Caption}' matches a source field name, which Excel does not allow. Use a caption such as 'Sum of {valueField.Caption}'.");
            }

            if (!captions.Add(valueField.Caption))
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' has more than one value field captioned '{valueField.Caption}'.");
            }
        }

        // Report filters occupy the rows above the table, followed by one blank row.
        if (pivotTable.FilterFields.Count > 0 && pivotTable.StartCell.Row < pivotTable.FilterFields.Count + 2)
        {
            throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' has {pivotTable.FilterFields.Count} filter field(s) and needs to start at row {pivotTable.FilterFields.Count + 2} or lower so the filters fit above it.");
        }
    }

    private static void ValidatePivotFields(ExcelPivotTableSpec pivotTable, IReadOnlyList<string> headers)
    {
        foreach (var fieldName in pivotTable.RowFields.Concat(pivotTable.ColumnFields).Concat(pivotTable.FilterFields).Concat(pivotTable.ValueFields.Select(field => field.FieldName)))
        {
            if (!headers.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Pivot table '{pivotTable.Name}' field '{fieldName}' was not found in the source range headers.");
            }
        }
    }

    private static object? GetRenderedCellValue(SortedDictionary<int, SortedDictionary<int, Cell>> rows, int rowNumber, int columnNumber)
    {
        if (!rows.TryGetValue(rowNumber, out var row) || !row.TryGetValue(columnNumber, out var cell))
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.Boolean)
        {
            return cell.CellValue?.Text == "1";
        }

        if (cell.DataType?.Value == CellValues.Number && double.TryParse(cell.CellValue?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.Text?.Text ?? string.Empty;
        }

        return cell.CellValue?.Text;
    }

    private static string FormatPivotCacheKey(object? value)
    {
        // Values are compared by kind and text together, so the number 1 and the text "1" are different items.
        return value is null ? "<blank>" : $"{GetPivotValueKind(value)}:{Convert.ToString(value, CultureInfo.InvariantCulture)}";
    }

    private static string GetPivotValueKind(object? value)
    {
        return value switch
        {
            null => "blank",
            bool => "boolean",
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal => "number",
            _ => "string"
        };
    }

    private static PivotTableDefinition CreatePivotTableDefinition(
        ExcelPivotTableSpec pivotTable,
        PivotSourceData sourceData,
        uint cacheId)
    {
        var fieldNames = sourceData.Headers;
        var renderedValueFields = pivotTable.ValueFields;
        var valueCount = renderedValueFields.Count;
        var hasColumnFields = pivotTable.ColumnFields.Count > 0;
        var rowItemCount = GetAxisItemCount(sourceData, pivotTable.RowFields);
        var columnItemCount = GetAxisItemCount(sourceData, pivotTable.ColumnFields);

        // With more than one value field Excel adds a "Values" pseudo field (index -2) to the
        // column axis. It is an extra header level: Q1 -> [Sales, Average Cost], Q2 -> [...].
        var addValuesToColumns = valueCount > 1;
        var columnLevels = pivotTable.ColumnFields.Count + (addValuesToColumns ? 1 : 0);

        // Each column level is a header row stacked above the data. Without column fields the
        // header is a single row (or, with several values, the row of value captions).
        var firstHeaderRow = hasColumnFields || !addValuesToColumns ? 1U : 0U;
        var firstDataRow = hasColumnFields ? 1U + (uint)columnLevels : 1U;
        // The pivot table's area is worked out here, since the location element must give its full extent.
        var showsGrandTotalRow = pivotTable.ColumnGrandTotals && pivotTable.RowFields.Count > 0;
        var outputRows = (int)firstDataRow + Math.Max(1, rowItemCount) + (showsGrandTotalRow ? 1 : 0);

        // Every column item is repeated for each value field, plus one grand total column per value.
        var valueColumnCount = Math.Max(1, columnItemCount) * valueCount
            + (hasColumnFields && pivotTable.RowGrandTotals ? valueCount : 0);
        var outputColumns = Math.Max(1, pivotTable.RowFields.Count) + valueColumnCount;
        var endCell = pivotTable.StartCell.Offset(outputRows - 1, outputColumns - 1);
        var definition = new PivotTableDefinition
        {
            Name = ExcelNames.SanitizePivotName(pivotTable.Name),
            CacheId = cacheId,
            DataCaption = "Values",
            ApplyNumberFormats = false,
            ApplyBorderFormats = false,
            ApplyFontFormats = false,
            ApplyPatternFormats = false,
            ApplyAlignmentFormats = false,
            ApplyWidthHeightFormats = true,
            RowGrandTotals = pivotTable.RowGrandTotals,
            ColumnGrandTotals = pivotTable.ColumnGrandTotals,
            UseAutoFormatting = true,
            PreserveFormatting = true,
            ItemPrintTitles = true,
            CreatedVersion = 8,
            UpdatedVersion = 8,
            MinRefreshableVersion = 3
        };

        var location = new Location
        {
            Reference = $"{pivotTable.StartCell.ToReference()}:{endCell.ToReference()}",
            FirstHeaderRow = firstHeaderRow,
            FirstDataRow = firstDataRow,
            FirstDataColumn = 1U
        };
        if (pivotTable.FilterFields.Count > 0)
        {
            // The filter cells sit above the table: one row per filter, laid out in a single column.
            location.RowPageCount = (uint)pivotTable.FilterFields.Count;
            location.ColumnsPerPage = 1U;
        }

        definition.Append(location);
        definition.Append(CreatePivotFields(pivotTable, sourceData, renderedValueFields));

        if (pivotTable.RowFields.Count > 0)
        {
            definition.Append(CreateAxisFields<RowFields>(pivotTable.RowFields, fieldNames));
            definition.Append(CreateRowItems(sourceData, pivotTable.RowFields, pivotTable.ColumnGrandTotals));
        }

        if (hasColumnFields || addValuesToColumns)
        {
            definition.Append(CreateAxisFields<ColumnFields>(pivotTable.ColumnFields, fieldNames, includeValuesField: addValuesToColumns));
            definition.Append(CreateColumnItems(sourceData, pivotTable.ColumnFields, valueCount, includeGrandTotal: hasColumnFields && pivotTable.RowGrandTotals));
        }

        if (pivotTable.FilterFields.Count > 0)
        {
            definition.Append(CreatePageFields(pivotTable.FilterFields, fieldNames));
        }

        definition.Append(CreateDataFields(renderedValueFields, fieldNames));
        definition.Append(new PivotTableStyle
        {
            Name = pivotTable.StyleName,
            ShowRowHeaders = true,
            ShowColumnHeaders = true,
            ShowRowStripes = false,
            ShowColumnStripes = false,
            ShowLastColumn = true
        });

        return definition;
    }

    private static PivotFields CreatePivotFields(
        ExcelPivotTableSpec pivotTable,
        PivotSourceData sourceData,
        IReadOnlyList<ExcelPivotValueFieldSpec> renderedValueFields)
    {
        var fieldNames = sourceData.Headers;
        var pivotFields = new PivotFields { Count = (uint)fieldNames.Count };
        foreach (var fieldName in fieldNames)
        {
            var field = new PivotField { ShowAll = false };
            PivotTableAxisValues? axis = null;
            if (pivotTable.RowFields.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                axis = PivotTableAxisValues.AxisRow;
            }
            else if (pivotTable.ColumnFields.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                axis = PivotTableAxisValues.AxisColumn;
            }
            else if (pivotTable.FilterFields.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
            {
                axis = PivotTableAxisValues.AxisPage;
            }

            if (axis is not null)
            {
                // Row, column and filter fields all list their distinct values as items.
                // A filter starts on "(All)", which is the trailing default item.
                field.Axis = axis.Value;
                field.Append(CreatePivotFieldItems(GetDistinctFieldValues(sourceData, fieldName)));
            }

            // A source field can feed a value and still sit on an axis (for example a count of Region),
            // and several value fields can share one source field (a sum and an average of Amount).
            if (renderedValueFields.Any(value => string.Equals(value.FieldName, fieldName, StringComparison.OrdinalIgnoreCase)))
            {
                field.DataField = true;
            }

            pivotFields.Append(field);
        }

        return pivotFields;
    }

    private static T CreateAxisFields<T>(IReadOnlyList<string> fields, IReadOnlyList<string> fieldNames, bool includeValuesField = false)
        where T : OpenXmlCompositeElement, new()
    {
        var axisFields = new T();
        var count = fields.Count + (includeValuesField ? 1 : 0);
        axisFields.SetAttribute(new OpenXmlAttribute("count", null, count.ToString(CultureInfo.InvariantCulture)));
        foreach (var field in fields)
        {
            axisFields.Append(new Field { Index = FindPivotFieldIndex(fieldNames, field) });
        }

        if (includeValuesField)
        {
            // Index -2 is not a source field. It stands for the "Values" header that lists the value fields.
            axisFields.Append(new Field { Index = -2 });
        }

        return axisFields;
    }

    private static Items CreatePivotFieldItems(IReadOnlyList<object?> values)
    {
        var items = new Items { Count = (uint)(values.Count + 1) };
        for (var i = 0; i < values.Count; i++)
        {
            items.Append(new Item { Index = (uint)i });
        }

        items.Append(new Item { ItemType = ItemValues.Default });
        return items;
    }

    private static RowItems CreateRowItems(PivotSourceData sourceData, IReadOnlyList<string> fields, bool includeGrandTotal)
    {
        var items = new List<DocumentFormat.OpenXml.Spreadsheet.RowItem>();
        foreach (var key in GetDistinctAxisKeys(sourceData, fields))
        {
            items.Add(CreateAxisItem(key.Indexes));
        }

        if (includeGrandTotal)
        {
            items.Add(CreateAxisItem([], itemType: ItemValues.Grand));
        }

        var rowItems = new RowItems { Count = (uint)items.Count };
        rowItems.Append(items);
        return rowItems;
    }

    /// <summary>
    /// Builds the column items. With several value fields every column item is repeated once per value
    /// (Q1 Sales, Q1 Average Cost, Q2 Sales, ...) and each repeat carries its value index in the "i" attribute
    /// and as a last "x" entry, exactly as Excel writes them.
    /// </summary>
    private static ColumnItems CreateColumnItems(PivotSourceData sourceData, IReadOnlyList<string> fields, int valueCount, bool includeGrandTotal)
    {
        var items = new List<DocumentFormat.OpenXml.Spreadsheet.RowItem>();
        var multiValue = valueCount > 1;
        var keys = fields.Count == 0 ? [] : GetDistinctAxisKeys(sourceData, fields);

        if (keys.Count == 0)
        {
            // No column fields: the only columns are the value fields themselves.
            for (var valueIndex = 0; valueIndex < valueCount; valueIndex++)
            {
                items.Add(CreateAxisItem([], dataIndex: valueIndex));
            }
        }
        else
        {
            foreach (var key in keys)
            {
                if (!multiValue)
                {
                    items.Add(CreateAxisItem(key.Indexes));
                    continue;
                }

                for (var valueIndex = 0; valueIndex < valueCount; valueIndex++)
                {
                    items.Add(CreateAxisItem(key.Indexes, dataIndex: valueIndex));
                }
            }
        }

        if (includeGrandTotal)
        {
            for (var valueIndex = 0; valueIndex < (multiValue ? valueCount : 1); valueIndex++)
            {
                items.Add(CreateAxisItem([], itemType: ItemValues.Grand, dataIndex: multiValue ? valueIndex : null));
            }
        }

        var columnItems = new ColumnItems { Count = (uint)items.Count };
        columnItems.Append(items);
        return columnItems;
    }

    private static DocumentFormat.OpenXml.Spreadsheet.RowItem CreateAxisItem(
        IReadOnlyList<int> fieldItemIndexes,
        ItemValues? itemType = null,
        int? dataIndex = null)
    {
        var item = new DocumentFormat.OpenXml.Spreadsheet.RowItem();
        if (itemType is not null)
        {
            item.ItemType = itemType.Value;
        }

        // Excel leaves the attribute off for the first value field and writes it for the others.
        if (dataIndex is > 0)
        {
            item.Index = (uint)dataIndex.Value;
        }

        foreach (var index in fieldItemIndexes)
        {
            item.Append(new FieldItem { Val = (uint)index });
        }

        if (dataIndex is not null)
        {
            item.Append(new FieldItem { Val = (uint)dataIndex.Value });
        }
        else if (fieldItemIndexes.Count == 0)
        {
            // A grand total item still lists one (empty) level.
            item.Append(new FieldItem());
        }

        return item;
    }

    private static PageFields CreatePageFields(IReadOnlyList<string> fields, IReadOnlyList<string> fieldNames)
    {
        var pageFields = new PageFields { Count = (uint)fields.Count };
        foreach (var field in fields)
        {
            // Hierarchy -1 means an ordinary field rather than an OLAP hierarchy. No item is selected, so it shows "(All)".
            pageFields.Append(new PageField { Field = FindPivotFieldIndex(fieldNames, field), Hierarchy = -1 });
        }

        return pageFields;
    }

    private static DataFields CreateDataFields(IReadOnlyList<ExcelPivotValueFieldSpec> valueFields, IReadOnlyList<string> fieldNames)
    {
        var dataFields = new DataFields { Count = (uint)valueFields.Count };
        foreach (var valueField in valueFields)
        {
            dataFields.Append(new DataField
            {
                Name = valueField.Caption,
                Field = (uint)FindPivotFieldIndex(fieldNames, valueField.FieldName),
                Subtotal = ToOpenXmlFunction(valueField.Function),
                BaseField = 0,
                BaseItem = 0U
            });
        }

        return dataFields;
    }

    private static IReadOnlyList<string> ResolvePivotFieldNames(ExcelPivotTableSpec pivotTable)
    {
        return pivotTable.RowFields
            .Concat(pivotTable.ColumnFields)
            .Concat(pivotTable.FilterFields)
            .Concat(pivotTable.ValueFields.Select(field => field.FieldName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static int FindPivotFieldIndex(IReadOnlyList<string> fieldNames, string fieldName)
    {
        for (var i = 0; i < fieldNames.Count; i++)
        {
            if (string.Equals(fieldNames[i], fieldName, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        throw new InvalidOperationException($"Pivot field '{fieldName}' was not found in the pivot source field list.");
    }

    private static int GetAxisItemCount(PivotSourceData sourceData, IReadOnlyList<string> fields)
    {
        return fields.Count == 0 ? 0 : GetDistinctAxisKeys(sourceData, fields).Count;
    }

    private static IReadOnlyList<object?> GetDistinctFieldValues(PivotSourceData sourceData, string fieldName)
    {
        var fieldIndex = FindPivotFieldIndex(sourceData.Headers, fieldName);
        return sourceData.Rows
            .Select(row => fieldIndex < row.Count ? row[fieldIndex] : null)
            .DistinctBy(FormatPivotCacheKey)
            .ToList();
    }

    private static IReadOnlyList<PivotAxisKey> GetDistinctAxisKeys(PivotSourceData sourceData, IReadOnlyList<string> fields)
    {
        var fieldIndexes = fields.Select(field => FindPivotFieldIndex(sourceData.Headers, field)).ToList();
        var distinctFieldValues = fields.ToDictionary(field => field, field => GetDistinctFieldValues(sourceData, field), StringComparer.OrdinalIgnoreCase);
        var keys = new List<PivotAxisKey>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var row in sourceData.Rows)
        {
            var values = new List<object?>();
            var itemIndexes = new List<int>();
            for (var i = 0; i < fields.Count; i++)
            {
                var value = fieldIndexes[i] < row.Count ? row[fieldIndexes[i]] : null;
                values.Add(value);
                itemIndexes.Add(IndexOfPivotValue(distinctFieldValues[fields[i]], value));
            }

            var key = string.Join("|", values.Select(FormatPivotCacheKey));
            if (seen.Add(key))
            {
                keys.Add(new PivotAxisKey(key, itemIndexes));
            }
        }

        // Axis items must be grouped by their outer fields (all of Europe together, then Americas, ...),
        // in the same order as the field's items. Source-row order would split a group apart.
        return keys.OrderBy(key => key.Indexes, Comparer<IReadOnlyList<int>>.Create(CompareItemIndexes)).ToList();
    }

    private static int CompareItemIndexes(IReadOnlyList<int> left, IReadOnlyList<int> right)
    {
        for (var i = 0; i < Math.Min(left.Count, right.Count); i++)
        {
            var comparison = left[i].CompareTo(right[i]);
            if (comparison != 0)
            {
                return comparison;
            }
        }

        return left.Count.CompareTo(right.Count);
    }

    private static int IndexOfPivotValue(IReadOnlyList<object?> values, object? value)
    {
        var key = FormatPivotCacheKey(value);
        for (var i = 0; i < values.Count; i++)
        {
            if (string.Equals(FormatPivotCacheKey(values[i]), key, StringComparison.Ordinal))
            {
                return i;
            }
        }

        // A value that is not found falls back to the first item rather than failing.
        return 0;
    }

    private static DataConsolidateFunctionValues ToOpenXmlFunction(ExcelPivotAggregateFunction function)
    {
        return function switch
        {
            ExcelPivotAggregateFunction.Count => DataConsolidateFunctionValues.Count,
            ExcelPivotAggregateFunction.Average => DataConsolidateFunctionValues.Average,
            ExcelPivotAggregateFunction.Minimum => DataConsolidateFunctionValues.Minimum,
            ExcelPivotAggregateFunction.Maximum => DataConsolidateFunctionValues.Maximum,
            _ => DataConsolidateFunctionValues.Sum
        };
    }
}
