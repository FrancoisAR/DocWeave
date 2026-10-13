using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

internal static class ExcelWorkbookRenderer
{
    public static void Render(ExcelWorkbookSpec workbookSpec, Stream target)
    {
        using var buffer = new MemoryStream();
        RenderCore(workbookSpec, buffer);
        buffer.Position = 0;
        using var normalized = ExcelPackageNormalizer.NormalizePackage(buffer);
        normalized.Position = 0;
        normalized.CopyTo(target);
    }

    private static void RenderCore(ExcelWorkbookSpec workbookSpec, Stream target)
    {
        if (workbookSpec.Sheets.Count == 0)
        {
            throw new InvalidOperationException("At least one worksheet must be added before exporting an Excel workbook.");
        }

        using var document = SpreadsheetDocument.Create(target, SpreadsheetDocumentType.Workbook, autoSave: true);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        if (workbookSpec.Protection is not null)
        {
            workbookPart.Workbook.Append(ExcelSheetFeatures.CreateWorkbookProtection(workbookSpec.Protection));
        }

        var styleRegistry = new ExcelStyleRegistry();
        var nextPivotCacheId = 1U;
        var renderedSheets = new List<RenderedWorksheetContext>();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        for (var i = 0; i < workbookSpec.Sheets.Count; i++)
        {
            renderedSheets.Add(AddWorksheet(workbookPart, sheets, workbookSpec.Sheets[i], i + 1, styleRegistry));
        }

        var renderedSheetLookup = renderedSheets.ToDictionary(
            sheet => ExcelNames.SanitizeSheetName(sheet.SheetSpec.Name),
            sheet => sheet,
            StringComparer.OrdinalIgnoreCase);

        foreach (var renderedSheet in renderedSheets)
        {
            if (renderedSheet.SheetSpec.PivotTables.Count > 0)
            {
                ExcelPivotRenderer.AddPivotTables(workbookPart, renderedSheet, renderedSheetLookup, ref nextPivotCacheId);
                renderedSheet.WorksheetPart.Worksheet.Save();
            }
        }

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = styleRegistry.CreateStylesheet();
        stylesPart.Stylesheet.Save();
        workbookPart.Workbook.Save();
    }

    private static RenderedWorksheetContext AddWorksheet(
        WorkbookPart workbookPart,
        Sheets sheets,
        ExcelSheetSpec sheetSpec,
        int sheetId,
        ExcelStyleRegistry styleRegistry)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        if (sheetSpec.FreezePane is not null)
        {
            worksheetPart.Worksheet.InsertBefore(ExcelSheetFeatures.CreateSheetViews(sheetSpec.FreezePane), sheetData);
        }

        var rows = new SortedDictionary<int, SortedDictionary<int, Cell>>();
        foreach (var cell in sheetSpec.Cells)
        {
            ExcelCellWriter.SetCell(rows, sheetSpec, cell.Address.Row, cell.Address.Column, cell.Value, cell.Style, styleRegistry);
        }

        foreach (var table in sheetSpec.Tables)
        {
            ExcelTableRenderer.WriteTable(rows, sheetSpec, table, styleRegistry);
        }

        foreach (var repeater in sheetSpec.Repeaters)
        {
            ExcelTableRenderer.WriteRepeater(rows, sheetSpec, repeater, styleRegistry);
        }

        foreach (var rangeStyle in sheetSpec.RangeStyles)
        {
            ExcelCellWriter.ApplyRangeStyle(rows, sheetSpec, rangeStyle, styleRegistry);
        }

        if (sheetSpec.ColumnSettings.Count > 0)
        {
            worksheetPart.Worksheet.InsertBefore(CreateColumns(sheetSpec), sheetData);
        }

        foreach (var rowNumber in sheetSpec.RowSettings.Keys)
        {
            if (!rows.ContainsKey(rowNumber))
            {
                rows[rowNumber] = [];
            }
        }

        // OpenXML stores cells in rows, and rows in ascending order. Keeping a small in-memory
        // map here makes overlapping regions deterministic until we add validation diagnostics.
        foreach (var rowItem in rows)
        {
            var row = CreateRow(rowItem.Key, sheetSpec);
            foreach (var cell in rowItem.Value.Values)
            {
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        OpenXmlElement insertAfter = sheetData;
        if (sheetSpec.Protection is not null)
        {
            var protection = ExcelSheetFeatures.CreateSheetProtection(sheetSpec.Protection);
            worksheetPart.Worksheet.InsertAfter(protection, insertAfter);
            insertAfter = protection;
        }

        if (!string.IsNullOrWhiteSpace(sheetSpec.AutoFilterRange))
        {
            var autoFilter = new AutoFilter { Reference = sheetSpec.AutoFilterRange };
            worksheetPart.Worksheet.InsertAfter(autoFilter, insertAfter);
            insertAfter = autoFilter;
        }

        if (sheetSpec.MergeRanges.Count > 0)
        {
            var mergeCells = CreateMergeCells(sheetSpec);
            worksheetPart.Worksheet.InsertAfter(mergeCells, insertAfter);
            insertAfter = mergeCells;
        }

        if (sheetSpec.DataValidations.Count > 0)
        {
            worksheetPart.Worksheet.Append(ExcelSheetFeatures.CreateDataValidations(sheetSpec));
        }

        if (sheetSpec.Charts.Count > 0)
        {
            ExcelChartRenderer.AddCharts(worksheetPart, sheetSpec);
        }

        if (sheetSpec.Comments.Count > 0)
        {
            ExcelSheetFeatures.AddComments(worksheetPart, sheetSpec);
        }

        ExcelSheetFeatures.AddNativeTables(worksheetPart, sheetSpec);
        worksheetPart.Worksheet.Save();

        var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
        sheets.Append(new Sheet
        {
            Id = relationshipId,
            SheetId = (uint)sheetId,
            Name = ExcelNames.SanitizeSheetName(sheetSpec.Name)
        });

        return new RenderedWorksheetContext(sheetSpec, worksheetPart, rows);
    }

    private static Columns CreateColumns(ExcelSheetSpec sheetSpec)
    {
        var columns = new Columns();
        foreach (var item in sheetSpec.ColumnSettings.OrderBy(item => item.Key))
        {
            var settings = item.Value;
            columns.Append(new Column
            {
                Min = (uint)item.Key,
                Max = (uint)item.Key,
                Width = settings.Width,
                CustomWidth = settings.Width is not null,
                BestFit = settings.BestFit,
                Hidden = settings.Hidden
            });
        }

        return columns;
    }

    private static Row CreateRow(int rowNumber, ExcelSheetSpec sheetSpec)
    {
        var row = new Row { RowIndex = (uint)rowNumber };
        if (!sheetSpec.RowSettings.TryGetValue(rowNumber, out var settings))
        {
            return row;
        }

        row.Hidden = settings.Hidden;
        if (settings.Height is not null)
        {
            row.Height = settings.Height.Value;
            row.CustomHeight = true;
        }

        if (settings.AutoFit)
        {
            row.CustomHeight = false;
        }

        return row;
    }

    private static MergeCells CreateMergeCells(ExcelSheetSpec sheetSpec)
    {
        var mergeCells = new MergeCells();
        foreach (var range in sheetSpec.MergeRanges.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            mergeCells.Append(new MergeCell { Reference = range });
        }

        mergeCells.Count = (uint)mergeCells.ChildElements.Count;
        return mergeCells;
    }
}

/// <summary>
/// A worksheet after its cells have been written, kept so later passes (pivots reading source cells) can use it.
/// </summary>
internal sealed record RenderedWorksheetContext(
    ExcelSheetSpec SheetSpec,
    WorksheetPart WorksheetPart,
    SortedDictionary<int, SortedDictionary<int, Cell>> Rows);
