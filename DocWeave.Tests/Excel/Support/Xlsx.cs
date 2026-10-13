using System.Globalization;
using DocWeave.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Reads generated .xlsx files back for assertions: counts of parts and features, cell contents, and validity.
/// Every method takes the workbook bytes, so tests do not need to open the package themselves.
/// </summary>
internal static class Xlsx
{
    /// <summary>
    /// Fails the test when the workbook does not pass OpenXML schema validation, listing what is wrong.
    /// </summary>
    internal static void AssertValid(byte[] workbookBytes)
    {
        var errors = GetValidationErrors(workbookBytes);
        Assert.True(
            errors.Count == 0,
            $"The workbook has {errors.Count} OpenXML validation error(s):{Environment.NewLine}"
            + string.Join(Environment.NewLine, errors.Take(15).Select(error => $"  - {error}")));
    }

    /// <summary>
    /// Every OpenXML validation problem in the workbook, as "description (at path)".
    /// </summary>
    internal static IReadOnlyList<string> GetValidationErrors(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return new OpenXmlValidator()
            .Validate(document)
            .Select(error => $"{error.Description} (at {error.Path?.XPath})")
            .ToList();
    }

    internal static int CountOpenXmlValidationErrors(byte[] workbookBytes)
    {
        return GetValidationErrors(workbookBytes).Count;
    }

    internal static int CountWorksheets(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Count();
    }

    internal static int GetMaxUsedColumn(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Descendants<Cell>()
            .Select(cell => ParseCellReference(cell.CellReference!).column)
            .DefaultIfEmpty()
            .Max();
    }

    internal static int CountStyledCells(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        return worksheetPart.Worksheet.Descendants<Cell>().Count(cell => cell.StyleIndex?.Value > 0);
    }

    internal static int CountFormulaCells(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        return worksheetPart.Worksheet.Descendants<Cell>().Count(cell => cell.CellFormula is not null);
    }

    internal static string? GetCellFormula(byte[] workbookBytes, string sheetName, string cellReference)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == cellReference).CellFormula?.Text;
    }

    internal static string? GetCellText(byte[] workbookBytes, string sheetName, string cellReference)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        var cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference == cellReference);
        return cell.InlineString?.Text?.Text ?? cell.CellValue?.Text;
    }

    internal static int CountChartParts(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!
            .WorksheetParts
            .Sum(part => part.DrawingsPart?.ChartParts.Count() ?? 0);
    }

    internal static int CountChartPartsOfType<TChart>(byte[] workbookBytes)
        where TChart : OpenXmlElement
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.DrawingsPart?.ChartParts ?? [])
            .Count(part => part.ChartSpace.Descendants<TChart>().Any());
    }

    internal static int CountPivotTableParts(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.PivotTableParts.Count();
    }

    internal static int CountPivotCacheDefinitionParts(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.PivotTableCacheDefinitionParts.Count();
    }

    internal static PivotTableDefinition GetPivotTableDefinition(byte[] workbookBytes, string sheetName)
    {
        return GetPivotTableDefinitions(workbookBytes, sheetName).Single();
    }

    internal static IReadOnlyList<PivotTableDefinition> GetPivotTableDefinitions(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.PivotTableParts.Select(part => part.PivotTableDefinition).ToList();
    }

    /// <summary>
    /// Reads the "v" attribute of each child "x" of a row or column item. A missing attribute means index 0.
    /// The SDK exposes these children as different classes depending on the parent, so read them generically.
    /// </summary>
    internal static List<uint> GetItemValues(OpenXmlElement item)
    {
        return item.ChildElements
            .Select(child => child.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "v").Value)
            .Select(value => value is null ? 0U : uint.Parse(value, System.Globalization.CultureInfo.InvariantCulture))
            .ToList();
    }

    internal static ((int row, int column) start, (int row, int column) end) ParseReferenceRange(string range)
    {
        var parts = range.Split(':');
        return (ParseCellReference(parts[0]), ParseCellReference(parts[1]));
    }

    internal static (int row, int column) ParseCellReference(string cellReference)
    {
        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(cellReference.SkipWhile(char.IsLetter).ToArray());

        var column = 0;
        foreach (var character in letters.ToUpperInvariant())
        {
            column = column * 26 + (character - 'A' + 1);
        }

        return (int.Parse(digits), column);
    }

    internal static int CountMergeRanges(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<MergeCells>().Single().Elements<MergeCell>().Count();
    }

    internal static int CountWrappedOrRotatedCells(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        var wrappedOrRotatedStyleIndexes = stylesheet.CellFormats!
            .Elements<CellFormat>()
            .Select((format, index) => new { format, index })
            .Where(item => item.format.Alignment?.WrapText?.Value == true || item.format.Alignment?.TextRotation?.Value > 0)
            .Select(item => (uint)item.index)
            .ToHashSet();

        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.Worksheet.Descendants<Cell>())
            .Count(cell => cell.StyleIndex is not null && wrappedOrRotatedStyleIndexes.Contains(cell.StyleIndex.Value));
    }

    internal static int CountComments(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.WorksheetCommentsPart?.Comments.CommentList?.Elements<Comment>().Count() ?? 0;
    }

    internal static int CountHiddenColumns(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<Columns>().SelectMany(columns => columns.Elements<Column>()).Count(column => column.Hidden?.Value == true);
    }

    internal static int CountHiddenRows(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Descendants<Row>().Count(row => row.Hidden?.Value == true);
    }

    internal static int CountPatternFills(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet.Fills!
            .Elements<Fill>()
            .Select(fill => fill.PatternFill?.PatternType?.Value)
            .Count(pattern => pattern is not null && pattern != PatternValues.None && pattern != PatternValues.Solid && pattern != PatternValues.Gray125);
    }

    internal static int CountDataValidations(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<DataValidations>().Single().Elements<DataValidation>().Count();
    }

    internal static bool HasWorkbookProtection(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.Workbook.Elements<WorkbookProtection>().Any();
    }

    internal static bool HasSheetProtection(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<SheetProtection>().Any();
    }

    internal static int CountNativeTables(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.TableDefinitionParts.Count();
    }

    internal static int CountChartDataLabels(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.DrawingsPart?.ChartParts ?? [])
            .Sum(part => part.ChartSpace.Descendants<C.DataLabels>().Count());
    }

    internal static int CountChartAxisTitles(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.DrawingsPart?.ChartParts ?? [])
            .Sum(part => part.ChartSpace.Descendants<C.CategoryAxis>().SelectMany(axis => axis.Elements<C.Title>()).Count()
                + part.ChartSpace.Descendants<C.ValueAxis>().SelectMany(axis => axis.Elements<C.Title>()).Count());
    }

    internal static Pane GetFreezePane(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<SheetViews>().Single().Elements<SheetView>().Single().Elements<Pane>().Single();
    }

    internal static string? GetAutoFilterReference(byte[] workbookBytes, string sheetName)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = GetWorksheetPart(document, sheetName);
        return worksheetPart.Worksheet.Elements<AutoFilter>().Single().Reference?.Value;
    }

    internal static int CountFormattedCells(byte[] workbookBytes, ExcelNumberFormat format)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var targetNumberFormatId = format switch
        {
            ExcelNumberFormat.Number => 4U,
            ExcelNumberFormat.Currency => 164U,
            ExcelNumberFormat.Percent => 10U,
            ExcelNumberFormat.Date => 14U,
            ExcelNumberFormat.Text => 49U,
            _ => 0U
        };

        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        var matchingStyleIndexes = stylesheet.CellFormats!
            .Elements<CellFormat>()
            .Select((formatItem, index) => new { formatItem, index })
            .Where(item => item.formatItem.NumberFormatId?.Value == targetNumberFormatId)
            .Select(item => (uint)item.index)
            .ToHashSet();

        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.Worksheet.Descendants<Cell>())
            .Count(cell => cell.StyleIndex is not null && matchingStyleIndexes.Contains(cell.StyleIndex.Value));
    }

    internal static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    internal static C.LegendPositionValues? GetFirstChartLegendPosition(byte[] workbookBytes)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!
            .WorksheetParts
            .SelectMany(part => part.DrawingsPart?.ChartParts ?? [])
            .Select(part => part.ChartSpace.Descendants<C.LegendPosition>().FirstOrDefault()?.Val?.Value)
            .FirstOrDefault(value => value is not null);
    }
}
