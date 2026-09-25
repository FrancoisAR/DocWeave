# DocWeave Interface Technical Plan

> **Status: design document.** It plans interfaces (such as schema binding and `FromSchema`) that are not all in the library yet. The current API is in [getting-started.md](getting-started.md) and [pre-release-functionality-guide.md](pre-release-functionality-guide.md).

## Purpose

This document defines the interface plan for implementing DocWeave as a composable import/export library. The primary output target is Excel `.xlsx`, with CSV and other structured exports as first-class companions. The design also leaves room for later Word document and PowerPoint deck generation using the same source binding, schema, template, style, and validation concepts.

The central idea is:

```text
User fluent API / schema
    -> immutable specs
    -> validation
    -> data adapters
    -> format-specific renderers/readers
    -> file, stream, byte array, or object result
```

The public API should feel fluent and task-oriented. The internal API should be spec-driven and testable.

## Core Design Principles

- Keep simple exports simple.
- Treat each output area as an export region.
- Let one workbook contain many sheets and many regions per sheet.
- Allow mixed data sources in one export: `DataSet`, `DataTable`, `IEnumerable<T>`, dictionaries, dynamic objects, scalar values, custom adapters, and binary assets.
- Allow each region to have independent placement, style, columns, headers, formulas, totals, conditional formatting, empty-source behavior, and table/range names.
- Support grouped defaults with per-region overrides.
- Keep OpenXML details behind renderer interfaces.
- Keep schema and fluent APIs compiling to the same specs.
- Avoid coupling Excel-only concepts to CSV, Word, or PowerPoint renderers.
- Treat logging and audit as first-class cross-cutting concerns.
- Never log source row values, passwords, secrets, or generated file content by default.

## High-Level Interface Layers

```text
DocWeave
  Public fluent API
  Declarative schema API
  Specification model
  Validation
  Data adapters
  Layout engine
  Excel renderers
  CSV/structured renderers
  Import readers
  Template services
  Observability and audit services
```

## Public Export Entry Points

```csharp
public interface IExcelReportStart
{
    IExcelWorkbookBuilder ToFile(string path);
    IExcelWorkbookBuilder ToStream(Stream stream);
    IExcelWorkbookByteArrayBuilder ToByteArray();
    IExcelWorkbookBuilder UseTemplate(string templatePath);
    IExcelWorkbookBuilder UseTemplate(Stream templateStream);
    IExcelWorkbookBuilder UseTemplate(string templatePath, TemplateOptions options);
    IExcelWorkbookBuilder UseTemplate(Stream templateStream, TemplateOptions options);
}

public interface IExcelWorkbookBuilder
{
    IExcelWorkbookBuilder Properties(Action<IWorkbookPropertiesBuilder> configure);
    IExcelWorkbookBuilder PreserveMacros(bool preserve = true);
    IExcelWorkbookBuilder Encrypt(Action<IWorkbookEncryptionBuilder> configure);
    IExcelWorkbookBuilder Protect(Action<IWorkbookProtectionBuilder> configure);
    IExcelWorkbookBuilder Sheet(string name, Action<IExcelSheetBuilder> configure);
    IExcelWorkbookBuilder SheetsFrom(DataSet dataSet);
    IExcelWorkbookBuilder DataSet(DataSet dataSet, Action<IDataSetExportBuilder> configure);
    IExcelWorkbookBuilder WithSource(string name, object source);
    IExcelWorkbookBuilder WithValue(string name, object? value);
    IExcelWorkbookBuilder WithAsset(string name, Stream content, string contentType);
    WorkbookSpec Build();
    void Save();
}

public interface IExcelWorkbookByteArrayBuilder : IExcelWorkbookBuilder
{
    new byte[] Save();
}

public interface IWorkbookEncryptionBuilder
{
    IWorkbookEncryptionBuilder OpenPassword(string password);
    IWorkbookEncryptionBuilder ModifyPassword(string password);
    IWorkbookEncryptionBuilder Algorithm(ExcelEncryptionAlgorithm algorithm);
    IWorkbookEncryptionBuilder RequirePasswordToOpen();
}

public sealed record TemplateOptions(
    string? FileName = null,
    bool PreserveMacros = false);
```

Static facade:

```csharp
public static class ExcelReport
{
    public static IExcelReportStart Create();
    public static IExcelWorkbookBuilder FromSchema(LayoutSchema schema);
    public static IExcelWorkbookBuilder FromSchema(string json);
}
```

## Sheet Builder Interfaces

```csharp
public interface IExcelSheetBuilder
{
    IExcelSheetBuilder StartAt(string cell);
    IExcelSheetBuilder StartAt(int row, int column);
    IExcelSheetBuilder TabColor(string hex);
    IExcelSheetBuilder Hide();
    IExcelSheetBuilder VeryHide();
    IExcelSheetBuilder HideGridlines();
    IExcelSheetBuilder Freeze(int row, int column = 0);
    IExcelSheetBuilder PageSetup(Action<IPageSetupBuilder> configure);
    IExcelSheetBuilder PrintArea(string range);

    IExcelSheetBuilder From<T>(IEnumerable<T> rows);
    IExcelSheetBuilder From(DataTable table);
    IExcelSheetBuilder From(IRowSource source);

    IExcelSheetBuilder Sources(Action<IExportRegionCollectionBuilder> configure);
    IExcelSheetBuilder Group(string name, Action<IDataGroupBuilder> configure);
    IExcelSheetBuilder RepeatBlocks<T>(IEnumerable<T> items, Action<IRepeatingBlockBuilder<T>> configure);
    IExcelSheetBuilder PivotTable(string name, Action<IPivotTableBuilder> configure);
    IExcelSheetBuilder Chart(string name, Action<IChartBuilder> configure);
    IExcelSheetBuilder Image(string name, Action<IImageBuilder> configure);
    IExcelSheetBuilder Diagram(string name, Action<IDiagramBuilder> configure);

    IExcelSheetBuilder Validation(Action<IDataValidationBuilder> configure);
    IExcelSheetBuilder Hyperlinks(Action<IHyperlinkBuilder> configure);
    IExcelSheetBuilder Comments(Action<ICommentBuilder> configure);
    IExcelSheetBuilder Protect(Action<ISheetProtectionBuilder> configure);
}
```

## Export Region Interfaces

An export region is the core layout unit. Tables, summary ranges, repeated blocks, pivots, charts, images, diagrams, and template-bound areas should all compile into `ExportRegionSpec`.

```csharp
public interface IExportRegionCollectionBuilder
{
    IExportRegionBuilder Add<T>(string name, IEnumerable<T> rows);
    IExportRegionBuilder Add(string name, DataTable table);
    IExportRegionBuilder Add(string name, IRowSource source);
    IExportRegionBuilder AddValue(string name, object? value);
    IExportRegionBuilder AddKeyValueBlock(string name, object values);
    IExportRegionCollectionBuilder Layout(Action<IRegionLayoutBuilder> configure);
}

public interface IExportRegionBuilder
{
    IExportRegionBuilder At(string cell);
    IExportRegionBuilder After(string regionName, int spacingRows = 1);
    IExportRegionBuilder RightOf(string regionName, int spacingColumns = 1);
    IExportRegionBuilder AsTable(string tableName);
    IExportRegionBuilder AsNamedRange(string rangeName);
    IExportRegionBuilder AsKeyValueBlock();
    IExportRegionBuilder Headers(Action<IHeaderBuilder> configure);
    IExportRegionBuilder Columns(Action<IColumnBuilder> configure);
    IExportRegionBuilder Style(string styleName);
    IExportRegionBuilder Style(Action<IStyleBuilder> configure);
    IExportRegionBuilder Totals(Action<ITotalsBuilder> configure);
    IExportRegionBuilder Formulas(Action<IFormulaBuilder> configure);
    IExportRegionBuilder ConditionalFormatting(Action<IConditionalFormattingBuilder> configure);
    IExportRegionBuilder EmptySource(Action<IEmptySourceBuilder> configure);
    IExportRegionBuilder PageBreakAfter();
    IExportRegionBuilder KeepTogether();
}

public interface IRegionLayoutBuilder
{
    IRegionLayoutBuilder StackVertically(int spacingRows = 1);
    IRegionLayoutBuilder StackHorizontally(int spacingColumns = 1);
    IRegionLayoutBuilder Grid(int columns, int rowSpacing = 1, int columnSpacing = 1);
    IRegionLayoutBuilder AllowOverlap();
}
```

## DataSet Export Interfaces

```csharp
public interface IDataSetExportBuilder
{
    IDataSetExportBuilder AllTables();
    IDataSetTableExportBuilder Table(string tableName);
    IDataSetExportBuilder Ignore(string tableName);
    IDataSetExportBuilder UseRelations();
}

public interface IDataSetTableExportBuilder : IExportRegionBuilder
{
    IDataSetTableExportBuilder ToSheet(string sheetName);
}
```

Required behavior:

- Export all tables or selected tables.
- Use `DataTable.TableName` as default sheet/table name.
- Export each table to a separate sheet or combine tables on one sheet.
- Mix `DataSet` tables with other sources.
- Preserve `DataColumn` order and type metadata.
- Optionally use `DataRelation` for parent/child grouped output.

## Group Interfaces

Groups provide shared defaults for multiple regions.

```csharp
public interface IDataGroupBuilder
{
    IDataGroupBuilder StartAt(string cell);
    IDataGroupBuilder StackVertically(int spacingRows = 1);
    IDataGroupBuilder StackHorizontally(int spacingColumns = 1);
    IDataGroupBuilder Style(string styleName);
    IDataGroupBuilder Headers(Action<IHeaderBuilder> configure);
    IDataGroupBuilder Columns(Action<IColumnBuilder> configure);
    IDataGroupBuilder Totals(Action<ITotalsBuilder> configure);
    IDataGroupBuilder Formulas(Action<IFormulaBuilder> configure);

    IDataGroupBuilder Source<T>(string name, IEnumerable<T> rows, Action<IExportRegionBuilder>? configure = null);
    IDataGroupBuilder Source(string name, DataTable table, Action<IExportRegionBuilder>? configure = null);
    IDataGroupBuilder Source(string name, IRowSource source, Action<IExportRegionBuilder>? configure = null);
}
```

## Repeating Block Interfaces

Repeating blocks render each object as a rectangular region rather than a table row.

```csharp
public interface IRepeatingBlockBuilder<T>
{
    IRepeatingBlockBuilder<T> StartAt(string cell);
    IRepeatingBlockBuilder<T> PerRow(int count);
    IRepeatingBlockBuilder<T> BlockSize(int rows, int columns);
    IRepeatingBlockBuilder<T> Spacing(int rows = 0, int columns = 0);
    IRepeatingBlockBuilder<T> KeepTogether();
    IRepeatingBlockBuilder<T> PageBreakAfterRows(int blockRows);
    IRepeatingBlockBuilder<T> UseTemplateRange(string range);
    IRepeatingBlockBuilder<T> Bind(Action<IBlockBindingBuilder<T>> configure);
    IRepeatingBlockBuilder<T> ForEach(Action<IBlockInstanceBuilder<T>, T> configure);
}

public interface IBlockInstanceBuilder<T>
{
    IBlockCellBuilder Cell(string cellOrRange);
    IBlockInstanceBuilder<T> MiniTable<TChild>(
        string startCell,
        Func<T, IEnumerable<TChild>> children,
        Action<IExportRegionBuilder> configure);
}

public interface IBlockBindingBuilder<T>
{
    IBlockBindingBuilder<T> Value<TValue>(string placeholder, Func<T, TValue> value);
    IBlockBindingBuilder<T> Formula(string placeholder, string formula);
    IBlockBindingBuilder<T> MiniTable<TChild>(
        string placeholder,
        Func<T, IEnumerable<TChild>> children,
        Action<IExportRegionBuilder> configure);
}

public interface IBlockCellBuilder
{
    IBlockCellBuilder Value(object? value);
    IBlockCellBuilder Formula(string formula);
    IBlockCellBuilder Format(string format);
    IBlockCellBuilder Style(string styleName);
    IBlockCellBuilder Merge();
    IBlockCellBuilder MergeAcross(int columns);
    IBlockCellBuilder Wrap();
}
```

Placement formula:

```text
blockRow = itemIndex / perRow
blockColumn = itemIndex % perRow

targetRow = startRow + blockRow * (blockHeight + rowSpacing)
targetColumn = startColumn + blockColumn * (blockWidth + columnSpacing)
```

## Column, Header, Totals, Formula Interfaces

```csharp
public interface IColumnBuilder
{
    IColumnBuilder Include(string column);
    IColumnBuilder Include<T, TValue>(Expression<Func<T, TValue>> property);
    IColumnBuilder Ignore(string column);
    IColumnBuilder Rename(string column, string header);
    IColumnBuilder Order(params string[] columns);
    IColumnBuilder Width(string column, double width);
    IColumnBuilder AutoFit();
    IColumnBuilder Wrap(string column);
    IColumnBuilder Format<TValue>(string format);
    IColumnBuilder Format(string column, string format);
    IColumnBuilder Style(string column, string styleName);
    IColumnBuilder UseDisplayNameAttributes();
}

public interface IHeaderBuilder
{
    IHeaderBuilder Show();
    IHeaderBuilder Hide();
    IHeaderBuilder Freeze();
    IHeaderBuilder Filter();
    IHeaderBuilder Wrap();
    IHeaderBuilder Style(string styleName);
}

public interface ITotalsBuilder
{
    ITotalsBuilder Sum(string column);
    ITotalsBuilder Average(string column);
    ITotalsBuilder Count(string column);
    ITotalsBuilder Min(string column);
    ITotalsBuilder Max(string column);
    ITotalsBuilder Formula(string column, Func<FormulaContext, string> formula);
    ITotalsBuilder Label(string label);
    ITotalsBuilder Style(string styleName);
}

public interface IFormulaBuilder
{
    IFormulaBuilder Column(string column, string formula);
    IFormulaBuilder Column(string column, Func<int, string> formula);
    IFormulaBuilder RowTotal(string targetColumn, IReadOnlyList<string> sourceColumns);
    IFormulaBuilder Named(string name, string formula);
}
```

## Style Interfaces

```csharp
public interface IStyleBuilder
{
    IStyleBuilder Font(Action<IFontBuilder> configure);
    IStyleBuilder Fill(string hex);
    IStyleBuilder Foreground(string hex);
    IStyleBuilder Background(string hex);
    IStyleBuilder Alignment(Action<IAlignmentBuilder> configure);
    IStyleBuilder Border(Action<IBorderBuilder> configure);
    IStyleBuilder NumberFormat(string format);
}

public interface IFontBuilder
{
    IFontBuilder Name(string name);
    IFontBuilder Size(double points);
    IFontBuilder Bold();
    IFontBuilder Italic();
    IFontBuilder Underline();
    IFontBuilder Color(string hex);
}

public interface IAlignmentBuilder
{
    IAlignmentBuilder Horizontal(HorizontalAlignment alignment);
    IAlignmentBuilder Vertical(VerticalAlignment alignment);
    IAlignmentBuilder Wrap();
    IAlignmentBuilder Indent(int level);
}

public interface IBorderBuilder
{
    IBorderBuilder All(BorderStyle style, string? color = null);
    IBorderBuilder Top(BorderStyle style, string? color = null);
    IBorderBuilder Bottom(BorderStyle style, string? color = null);
    IBorderBuilder Left(BorderStyle style, string? color = null);
    IBorderBuilder Right(BorderStyle style, string? color = null);
}
```

Style resolution must be deterministic. For table exports, style layers should merge in this order:

1. Workbook/theme defaults.
2. Region/table style.
3. Header, row, alternate row, or total row style.
4. Column style.
5. Conditional formatting style.
6. Explicit cell/range override.

Later layers should override only the properties they define. For example, a highlighted column can add thick left/right borders while keeping the row font and alternating fill color from the row style.

## Conditional Formatting Interfaces

```csharp
public interface IConditionalFormattingBuilder
{
    IConditionalFormatRuleBuilder Range(string range);
    IConditionalFormatRuleBuilder Column(string column);
}

public interface IConditionalFormatRuleBuilder
{
    IConditionalFormatRuleBuilder GreaterThan(object value);
    IConditionalFormatRuleBuilder LessThan(object value);
    IConditionalFormatRuleBuilder Between(object min, object max);
    IConditionalFormatRuleBuilder EqualTo(object value);
    IConditionalFormatRuleBuilder TextContains(string value);
    IConditionalFormatRuleBuilder Top(int count);
    IConditionalFormatRuleBuilder Bottom(int count);
    IConditionalFormatRuleBuilder DuplicateValues();
    IConditionalFormatRuleBuilder UniqueValues();
    IConditionalFormatRuleBuilder ColorScale(string lowHex, string midHex, string highHex);
    IConditionalFormatRuleBuilder DataBar(string colorHex);
    IConditionalFormatRuleBuilder IconSet(string iconSetName);
    IConditionalFormatRuleBuilder Fill(string hex);
    IConditionalFormatRuleBuilder FontColor(string hex);
    IConditionalFormatRuleBuilder Style(string styleName);
}
```

## Pivot Table Interfaces

```csharp
public interface IPivotTableBuilder
{
    IPivotTableBuilder SourceTable(string tableName);
    IPivotTableBuilder SourceRange(string range);
    IPivotTableBuilder SourceNamedRange(string rangeName);
    IPivotTableBuilder StartAt(string cell);
    IPivotTableBuilder Rows(params string[] fields);
    IPivotTableBuilder Columns(params string[] fields);
    IPivotTableBuilder Filters(params string[] fields);
    IPivotTableBuilder Values(Action<IPivotValueBuilder> configure);
    IPivotTableBuilder ReportLayout(PivotReportLayout layout);
    IPivotTableBuilder ShowGrandTotals(bool rows = true, bool columns = true);
    IPivotTableBuilder RepeatItemLabels();
    IPivotTableBuilder Style(string styleName);
    IPivotTableBuilder RefreshOnOpen();
    IPivotTableBuilder PreserveCacheRecords(bool preserve = true);
}

public interface IPivotValueBuilder
{
    IPivotValueBuilder Sum(string field, string? caption = null);
    IPivotValueBuilder Count(string field, string? caption = null);
    IPivotValueBuilder Average(string field, string? caption = null);
    IPivotValueBuilder Min(string field, string? caption = null);
    IPivotValueBuilder Max(string field, string? caption = null);
    IPivotValueBuilder CountNumbers(string field, string? caption = null);
    IPivotValueBuilder StandardDeviation(string field, string? caption = null);
    IPivotValueBuilder Variance(string field, string? caption = null);
    IPivotValueBuilder Format(string field, string format);
}
```

## Chart And Visual Interfaces

```csharp
public interface IChartBuilder
{
    IChartBuilder SourceTable(string tableName);
    IChartBuilder SourceRange(string range);
    IChartBuilder SourceNamedRange(string rangeName);
    IChartBuilder SourcePivot(string pivotTableName);
    IChartBuilder Type(ChartType type);
    IChartBuilder Category(string fieldOrRange);
    IChartBuilder Series(string fieldOrRange, string? caption = null);
    IChartBuilder SecondaryAxis(string seriesName);
    IChartBuilder At(string cell);
    IChartBuilder Size(int width, int height);
    IChartBuilder Title(string title);
    IChartBuilder AxisTitles(string? horizontal = null, string? vertical = null);
    IChartBuilder Legend(ChartLegendPosition position);
    IChartBuilder DataLabels(Action<IChartDataLabelBuilder> configure);
    IChartBuilder Style(string styleName);
}

public interface IImageBuilder
{
    IImageBuilder FromFile(string path);
    IImageBuilder FromStream(Stream stream, string contentType);
    IImageBuilder FromBytes(byte[] bytes, string contentType);
    IImageBuilder FromAsset(string assetName);
    IImageBuilder At(string cell);
    IImageBuilder Size(int width, int height);
    IImageBuilder PreserveAspectRatio();
    IImageBuilder AltText(string text);
}

public interface IDiagramBuilder
{
    IDiagramBuilder At(string cell);
    IDiagramBuilder Shape(string name, Action<IShapeBuilder> configure);
    IDiagramBuilder Connector(string fromShape, string toShape, ConnectorType type);
}

public interface IShapeBuilder
{
    IShapeBuilder Rectangle();
    IShapeBuilder RoundedRectangle();
    IShapeBuilder TextBox();
    IShapeBuilder Callout();
    IShapeBuilder Text(string text);
    IShapeBuilder Size(int width, int height);
    IShapeBuilder At(string cell);
    IShapeBuilder RightOf(string shapeName, int spacing = 20);
    IShapeBuilder Below(string shapeName, int spacing = 20);
    IShapeBuilder Style(string styleName);
}

public interface ISparklineBuilder
{
    ISparklineBuilder Line(string name, string sourceRange, string targetCell);
    ISparklineBuilder Column(string name, string sourceRange, string targetCell);
    ISparklineBuilder WinLoss(string name, string sourceRange, string targetCell);
    ISparklineBuilder ForEachRow(string sourceColumns, string targetColumn);
}
```

Slicers and timelines should be separate later-phase interfaces because OpenXML compatibility is more complex:

```csharp
public interface ISlicerBuilder
{
    ISlicerBuilder ForTable(string tableName);
    ISlicerBuilder ForPivot(string pivotTableName);
    ISlicerBuilder Field(string fieldName);
    ISlicerBuilder At(string cell);
}
```

## Hyperlinks, Comments, Validation, Protection

```csharp
public interface IHyperlinkBuilder
{
    IHyperlinkTargetBuilder Cell(string cell);
    IHyperlinkTargetBuilder Column(string column);
}

public interface IHyperlinkTargetBuilder
{
    IHyperlinkTargetBuilder ToUrl(string url);
    IHyperlinkTargetBuilder ToFilePath();
    IHyperlinkTargetBuilder ToSheetRange(string sheetName, string range);
}

public interface ICommentBuilder
{
    ICommentBuilder Cell(string cell, string text, string? author = null);
}

public interface IDataValidationBuilder
{
    IDataValidationBuilder Column(string column, Action<IValidationRuleBuilder> configure);
    IDataValidationBuilder Range(string range, Action<IValidationRuleBuilder> configure);
}

public interface IValidationRuleBuilder
{
    IValidationRuleBuilder List(params string[] values);
    IValidationRuleBuilder WholeNumberBetween(int min, int max);
    IValidationRuleBuilder DecimalBetween(decimal min, decimal max);
    IValidationRuleBuilder DateBetween(DateTime min, DateTime max);
    IValidationRuleBuilder TextLengthBetween(int min, int max);
    IValidationRuleBuilder CustomFormula(string formula);
    IValidationRuleBuilder InputMessage(string title, string message);
    IValidationRuleBuilder ErrorMessage(string title, string message);
}

public interface ISheetProtectionBuilder
{
    ISheetProtectionBuilder ProtectSheet(string? password = null, bool allowFiltering = false, bool allowSorting = false);
    ISheetProtectionBuilder Password(string password);
    ISheetProtectionBuilder UnlockRange(string rangeNameOrAddress);
    ISheetProtectionBuilder HideFormulas(string rangeNameOrAddress);
    ISheetProtectionBuilder AllowFormattingCells();
    ISheetProtectionBuilder AllowInsertingRows();
    ISheetProtectionBuilder AllowDeletingRows();
    ISheetProtectionBuilder AllowSelectingUnlockedCells();
}

public interface IWorkbookProtectionBuilder
{
    IWorkbookProtectionBuilder ProtectWorkbook(string? password = null);
    IWorkbookProtectionBuilder ProtectStructure(string? password = null);
    IWorkbookProtectionBuilder Password(string password);
}
```

Protection and encryption must be treated separately:

- Sheet protection controls editing behavior for a worksheet.
- Workbook protection controls workbook structure, such as adding/removing sheets.
- File encryption controls whether a user can open the workbook package.
- Sheet-level confidentiality should be achieved by splitting sensitive data into separate encrypted files, or by using hidden/very-hidden sheets combined with full workbook encryption.

## CSV And Structured Export Interfaces

```csharp
public interface IStructuredDataExportStart
{
    IStructuredDataExportBuilder From<T>(IEnumerable<T> rows);
    IStructuredDataExportBuilder From(DataTable table);
    IStructuredDataExportBuilder From(IRowSource source);
}

public interface IStructuredDataExportBuilder
{
    IStructuredDataExportBuilder ToFile(string path);
    IStructuredDataExportBuilder ToStream(Stream stream);
    IStructuredDataExportBuilder ToTextWriter(TextWriter writer);
    IStructuredDataExportBuilder AsCsv();
    IStructuredDataExportBuilder AsDelimited(string delimiter);
    IStructuredDataExportBuilder AsFixedWidth(Action<IFixedWidthBuilder> configure);
    IStructuredDataExportBuilder AsJsonArray();
    IStructuredDataExportBuilder AsJsonLines();
    IStructuredDataExportBuilder AsXmlRecords();
    IStructuredDataExportBuilder AsHtmlTable();
    IStructuredDataExportBuilder AsMarkdownTable();
    IStructuredDataExportBuilder Encoding(Encoding encoding);
    IStructuredDataExportBuilder Headers(bool include = true);
    IStructuredDataExportBuilder Columns(Action<IColumnBuilder> configure);
    void Save();
}

public static class CsvWriter
{
    public static IStructuredDataExportStart Create();
}
```

CSV should support delimiter, quote, escape, encoding, header/no-header, null handling, culture, converters, malformed row policy, and streaming.

## Import Interfaces

```csharp
public interface IExcelReaderStart
{
    IExcelReadBuilder FromFile(string path, string? password = null);
    IExcelReadBuilder FromNetworkPath(string path, string? password = null);
    IExcelReadBuilder FromStream(Stream stream);
    IExcelReadBuilder FromStream(string fileName, Stream stream, string? password = null);
    IExcelReadBuilder FromBytes(byte[] bytes);
    IExcelReadBuilder FromBytes(string fileName, byte[] bytes, string? password = null);
    IExcelReadBuilder FromUpload(string fileName, Stream stream, string? password = null);
    IExcelReadBuilder FromUri(Uri uri, IImportFileResolver resolver, string? password = null);
}

public interface IExcelReadBuilder
{
    IExcelReadBuilder Label(string label);
    IExcelReadBuilder Password(string password);
    IExcelReadBuilder PasswordProvider(IImportPasswordProvider passwordProvider);
    IExcelReadBuilder WorkbookPassword(string password);
    IExcelSheetReadBuilder Sheet(string name);
    IExcelSheetReadBuilder Sheet(int index);
    IReadOnlyList<string> SheetNames();
}

public interface IExcelSheetReadBuilder
{
    IExcelSheetReadBuilder Password(string password);
    IExcelSheetReadBuilder UseHeaderRow(int row);
    IExcelSheetReadBuilder Range(string range);
    IExcelSheetReadBuilder Table(string tableName);
    IExcelSheetReadBuilder Columns(Action<IReadColumnBuilder> configure);
    IExcelSheetReadBuilder Conversion(Action<IConversionBuilder> configure);
    IExcelSheetReadBuilder OnError(ErrorMode mode);
    List<T> Read<T>();
    DataTable ReadDataTable();
    IReadOnlyList<IDictionary<string, object?>> ReadDictionaries();
    IReadOnlyList<RowValueSet> ReadRows();
    ImportResult<T> Validate<T>();
}

public interface IReadColumnBuilder
{
    IReadColumnBuilder Include(string source);
    IReadColumnBuilder Ignore(string source);
    IReadColumnBuilder IgnoreHidden();
    IReadColumnBuilder Rename(string source, string target);
    IReadColumnBuilder Type(string column, Type type);
    IReadColumnBuilder Nullable(string column, bool nullable = true);
    IReadColumnBuilder Default(string column, object? value);
    IReadColumnBuilder Add(string column, Func<ImportRowContext, object?> value);
    IReadColumnBuilder AddConstant(string column, object? value);
    IReadColumnBuilder AddFormula(string column, string expression);
    IReadColumnMapBuilder Map(string target);
}

public interface IReadColumnMapBuilder
{
    IReadColumnMapBuilder FromHeader(string header);
    IReadColumnMapBuilder FromColumn(string column);
    IReadColumnMapBuilder FromIndex(int index);
    IReadColumnMapBuilder As(Type type);
    IReadColumnMapBuilder AsString();
    IReadColumnMapBuilder AsInt32();
    IReadColumnMapBuilder AsDecimal();
    IReadColumnMapBuilder AsDateTime();
    IReadColumnMapBuilder AsBoolean(string trueValue = "true", string falseValue = "false");
    IReadColumnMapBuilder Nullable(bool nullable = true);
    IReadColumnMapBuilder ConvertWith(Func<object?, object?> converter);
}

public interface IExcelMappingBuilder<T>
{
    IExcelMappingBuilder<T> Include<TValue>(Expression<Func<T, TValue>> property);
    IExcelMappingBuilder<T> Ignore<TValue>(Expression<Func<T, TValue>> property);
    IExcelMappingBuilder<T> Property<TValue>(Expression<Func<T, TValue>> property);
    IExcelMappingBuilder<T> FromHeader(string header);
    IExcelMappingBuilder<T> FromColumn(string column);
    IExcelMappingBuilder<T> FromIndex(int index);
    IExcelMappingBuilder<T> As(Type type);
    IExcelMappingBuilder<T> AsString();
    IExcelMappingBuilder<T> AsInt32();
    IExcelMappingBuilder<T> AsDecimal();
    IExcelMappingBuilder<T> AsDateTime();
    IExcelMappingBuilder<T> AsBoolean(string trueValue = "true", string falseValue = "false");
    IExcelMappingBuilder<T> Default(object? value);
    IExcelMappingBuilder<T> ConvertWith(Func<object?, object?> converter);
    IExcelMappingBuilder<T> AddDerived<TValue>(Expression<Func<T, TValue>> property, Func<ImportRowContext, TValue> value);
    IExcelMappingBuilder<T> UseDisplayNameAttributes();
}

public interface IImportFileResolver
{
    Task<ResolvedImportFile> ResolveAsync(Uri uri, CancellationToken cancellationToken = default);
}

public interface IImportPasswordProvider
{
    string? GetPassword(ImportPasswordRequest request);
}

public sealed record ImportPasswordRequest(
    string? SourceLabel,
    string? FileName,
    string? SheetName,
    ImportPasswordKind Kind);

public sealed record ResolvedImportFile(
    string FileName,
    Stream Content,
    string? ContentType = null,
    long? Length = null);

public sealed class ImportRowContext
{
    public string? SourceLabel { get; init; }
    public string? SheetName { get; init; }
    public int RowNumber { get; init; }
    public object? this[string column] => GetValue(column);
    public object? GetValue(string column);
    public string? String(string column);
    public int Int32(string column);
    public decimal Decimal(string column);
    public DateTime DateTime(string column);
    public bool Boolean(string column);
}

public interface IImportBatchBuilder
{
    IImportBatchBuilder AddExcel(string name, byte[] bytes, string? fileName = null, string? password = null);
    IImportBatchBuilder AddExcel(string name, Stream stream, string? fileName = null, string? password = null);
    IImportBatchBuilder AddExcelPath(string name, string path, string? password = null);
    IImportBatchBuilder AddExcelUri(string name, Uri uri, IImportFileResolver resolver, string? password = null);
    IImportBatchBuilder AddCsv(string name, byte[] bytes, string? fileName = null);
    IImportBatchBuilder AddCsv(string name, Stream stream, string? fileName = null);
    IImportBatchBuilder AddCsvPath(string name, string path);
    IImportBatchResult Read(Action<IImportBatchReadBuilder> configure);
}

public interface IImportBatchReadBuilder
{
    IExcelReadBuilder Source(string name);
}
```

CSV reader mirrors Excel reader but targets delimited data:

```csharp
public interface ICsvReadBuilder
{
    ICsvReadBuilder Label(string label);
    ICsvReadBuilder DelimitedBy(string delimiter);
    ICsvReadBuilder Quote(char quote);
    ICsvReadBuilder Escape(char escape);
    ICsvReadBuilder Encoding(Encoding encoding);
    ICsvReadBuilder HasHeaderRow(bool hasHeader = true);
    ICsvReadBuilder TrimFields();
    ICsvReadBuilder Columns(Action<IReadColumnBuilder> configure);
    ICsvReadBuilder Conversion(Action<IConversionBuilder> configure);
    List<T> Read<T>();
    DataTable ReadDataTable();
    IReadOnlyList<IDictionary<string, object?>> ReadDictionaries();
}
```

Import source handling should stay policy-neutral in the core library. The library can accept paths, network paths, streams, byte arrays, uploads, and resolver-returned streams, but authentication, URL allow-lists, virus scanning, file size limits, and network access policy should belong to the WebAPI/application layer.

Import column shaping should run before typed object materialization:

1. Read raw source rows.
2. Resolve header/column/index mappings.
3. Apply include/ignore rules.
4. Rename output columns.
5. Convert values to declared types.
6. Add derived columns.
7. Materialize to `DataTable`, dictionaries, raw rows, typed objects, or validation results.

## Data Adapter Interfaces

```csharp
public interface IRowSource
{
    IReadOnlyList<ColumnDefinition> Columns { get; }
    IEnumerable<RowValueSet> Rows { get; }
}

public interface IRowSourceFactory
{
    bool CanCreate(object source);
    IRowSource Create(object source, DataSourceOptions options);
}

public interface IDataSetSourceAdapter
{
    IReadOnlyList<DataTableSource> GetTables(DataSet dataSet, DataSetExportOptions options);
}

public sealed record ColumnDefinition(
    string Key,
    string Header,
    Type DataType,
    int? Ordinal = null,
    string? Format = null);

public sealed record RowValueSet(IReadOnlyDictionary<string, object?> Values);
```

Required adapters:

- `ObjectEnumerableAdapter<T>`
- `DataTableAdapter`
- `DataSetAdapter`
- `DataRowEnumerableAdapter`
- `DictionaryAdapter`
- `DynamicObjectAdapter`
- `ScalarObjectAdapter`
- `CsvRecordAdapter`
- custom `IRowSource`

## Specification Model

Specs should be immutable records where practical.

```csharp
public sealed record WorkbookSpec(
    OutputTargetSpec Output,
    IReadOnlyList<SheetSpec> Sheets,
    IReadOnlyDictionary<string, DataSourceSpec> Sources,
    IReadOnlyDictionary<string, AssetSpec> Assets,
    StyleCatalogSpec Styles,
    WorkbookEncryptionSpec? Encryption,
    WorkbookOptions Options);

public sealed record SheetSpec(
    string Name,
    SheetOptions Options,
    IReadOnlyList<ExportRegionSpec> Regions);

public sealed record ExportRegionSpec(
    string Name,
    ExportRegionKind Kind,
    DataSourceRef? Source,
    RegionPlacementSpec Placement,
    TableSpec? Table,
    ColumnSetSpec? Columns,
    HeaderSpec? Header,
    StyleRef? Style,
    FormulaSetSpec? Formulas,
    TotalSetSpec? Totals,
    ConditionalFormatSetSpec? ConditionalFormatting,
    EmptySourceSpec? EmptySource,
    object? DetailSpec);
```

Specialized detail specs:

- `TableSpec`
- `RepeatingBlockSpec`
- `BlockTemplateSpec`
- `PivotTableSpec`
- `PivotFieldSpec`
- `ChartSpec`
- `DrawingSpec`
- `ImageSpec`
- `SparklineSpec`
- `SlicerSpec`
- `ValidationSpec`
- `HyperlinkSpec`
- `CommentSpec`
- `ProtectionSpec`
- `WorkbookEncryptionSpec`
- `ImportPasswordSpec`
- `TemplateBindingSpec`

## Schema Interfaces

```csharp
public interface ILayoutSchemaLoader
{
    LayoutSchema LoadJson(string json);
    LayoutSchema LoadYaml(string yaml);
    LayoutSchema LoadXml(string xml);
}

public interface ISchemaCompiler
{
    WorkbookSpec Compile(LayoutSchema schema, SchemaBindingContext bindings);
}

public interface ISchemaBindingContext
{
    ISchemaBindingContext WithSource(string name, object source);
    ISchemaBindingContext WithDataSet(string name, DataSet dataSet);
    ISchemaBindingContext WithValue(string name, object? value);
    ISchemaBindingContext WithAsset(string name, Stream content, string contentType);
    ISchemaBindingContext WithSecret(string name, string secret);
}
```

Schema source names should support:

- `activeDirectory`
- `transactions`
- `financialData.Costs`
- `financialData.Income`
- `financialData.Regions`

## Validation Interfaces

```csharp
public interface ISpecValidator<TSpec>
{
    ValidationResult Validate(TSpec spec);
}

public sealed record ValidationResult(
    bool IsValid,
    IReadOnlyList<ValidationIssue> Issues);

public sealed record ValidationIssue(
    ValidationSeverity Severity,
    string Code,
    string Message,
    string? Path = null);
```

Validation must check:

- Missing sources.
- Duplicate sheet/table/range names.
- Invalid sheet names.
- Invalid cell/range references.
- Invalid formulas and placeholders.
- Unsupported formats.
- Region collisions.
- Missing style references.
- Missing template ranges.
- Pivot source existence.
- Chart source existence.
- Unsupported chart/pivot combinations.
- Missing assets.

## Renderer Interfaces

```csharp
public interface IWorkbookRenderer
{
    Task RenderAsync(WorkbookSpec workbook, Stream output, OperationControl? control = null, CancellationToken cancellationToken = default);
}

public interface IExcelRenderer
{
    Task RenderAsync(WorkbookSpec workbook, ExcelRenderContext context, CancellationToken cancellationToken = default);
}

public interface IRegionRenderer
{
    bool CanRender(ExportRegionKind kind);
    RenderedRegion Render(ExportRegionSpec region, ExcelRenderContext context);
}

public interface ITableRenderer : IRegionRenderer { }
public interface IRepeatingBlockRenderer : IRegionRenderer { }
public interface IPivotTableRenderer : IRegionRenderer { }
public interface IChartRenderer : IRegionRenderer { }
public interface IDrawingRenderer : IRegionRenderer { }
public interface IImageRenderer : IRegionRenderer { }
```

Supporting Excel services:

```csharp
public interface IStyleRegistry
{
    uint GetOrCreateStyle(StyleSpec style);
}

public interface ISharedStringService
{
    int GetOrAdd(string value);
}

public interface ICellReferenceService
{
    CellReference Parse(string address);
    string ToAddress(int row, int column);
    string Offset(string address, int rows, int columns);
}

public interface IRegionLayoutEngine
{
    IReadOnlyList<ResolvedRegionPlacement> Resolve(SheetSpec sheet);
}

public interface IFormulaTranslator
{
    string Translate(string formula, FormulaContext context);
}

public interface ITemplateRangeCopier
{
    CopiedRange CopyRange(string sourceRange, string targetTopLeftCell, ExcelRenderContext context);
}

public interface IWorkbookEncryptionService
{
    Stream Encrypt(Stream unencryptedWorkbook, WorkbookEncryptionSpec encryption);
    Stream Decrypt(Stream encryptedWorkbook, ImportPasswordSpec password);
}

public interface IWorkbookPasswordService
{
    ImportPasswordSpec ResolvePassword(ImportPasswordRequest request);
}
```

Excel rendering order:

1. Validate `WorkbookSpec`.
2. Create or copy workbook/template.
3. Build style registry.
4. Create worksheets.
5. Resolve region layout per sheet.
6. Render table/data regions.
7. Render repeating blocks and template ranges.
8. Render pivot caches and pivot table definitions.
9. Render charts, images, diagrams, shapes, and anchors.
10. Apply formulas, filters, freezes, conditional formatting, validation, hyperlinks, comments, widths, merges, protection.
11. Save the unencrypted workbook package.
12. Apply full workbook encryption if `WorkbookEncryptionSpec` is present.

Encryption implementation note:

- Sheet/workbook protection is written as OpenXML workbook metadata.
- Full workbook encryption should be applied after the `.xlsx` package has been rendered.
- The encryption service should be optional/pluggable because full Office file encryption may need a specialized implementation or dependency.
- Passwords should not be logged, embedded in schema files, or included in validation output.

## Observability And Audit Interfaces

Logging and auditing should be separate from rendering logic. Renderers, validators, translators, and import readers should emit structured events through a small abstraction. The host application can connect those events to `ILogger`, OpenTelemetry, a database audit table, a message bus, or a SIEM platform.

```csharp
public interface IDocWeaveAuditSink
{
    ValueTask WriteAsync(DocWeaveAuditEvent auditEvent, CancellationToken cancellationToken = default);
}

public interface IDocWeaveOperationLogger
{
    IDisposable BeginOperation(DocWeaveOperationContext context);
    void Debug(string message, IReadOnlyDictionary<string, object?>? properties = null);
    void Information(string message, IReadOnlyDictionary<string, object?>? properties = null);
    void Warning(string message, IReadOnlyDictionary<string, object?>? properties = null);
    void Error(Exception exception, string message, IReadOnlyDictionary<string, object?>? properties = null);
}

public interface IDocWeaveTelemetry
{
    IDisposable StartActivity(string name, IReadOnlyDictionary<string, object?>? tags = null);
    void RecordMetric(string name, double value, IReadOnlyDictionary<string, object?>? tags = null);
}

public interface IDocWeaveProgressSink
{
    ValueTask ReportAsync(DocWeaveProgressEvent progress, CancellationToken cancellationToken = default);
}

public sealed record OperationControl(
    DocWeaveOperationContext Context,
    IDocWeaveProgressSink? ProgressSink = null,
    IDocWeaveAuditSink? AuditSink = null,
    IDocWeaveOperationLogger? Logger = null,
    IDocWeaveTelemetry? Telemetry = null);

public interface IAuditRedactionPolicy
{
    object? Redact(string propertyName, object? value);
    bool ShouldInclude(string propertyName, object? value);
}
```

Core audit records:

```csharp
public sealed record DocWeaveOperationContext(
    string OperationId,
    string OperationType,
    string? TemplateId,
    string? SchemaVersion,
    string? RequestedBy,
    string? CorrelationId,
    DateTimeOffset StartedAt);

public sealed record DocWeaveAuditEvent(
    string OperationId,
    string EventType,
    DateTimeOffset Timestamp,
    AuditSeverity Severity,
    IReadOnlyDictionary<string, object?> Properties);

public sealed record DocWeaveProgressEvent(
    string OperationId,
    string Stage,
    int? Percent,
    string? Message,
    IReadOnlyDictionary<string, object?> Properties);
```

Suggested audit event types:

- `ExportRequested`
- `ImportRequested`
- `SchemaValidated`
- `FluentTranslated`
- `SchemaTranslated`
- `SourcesBound`
- `LayoutResolved`
- `SheetRendered`
- `RegionRendered`
- `RowsRendered`
- `RowsImported`
- `ProgressReported`
- `OperationCancelled`
- `WorkbookEncrypted`
- `WorkbookProtected`
- `ValidationFailed`
- `ExportCompleted`
- `ImportCompleted`
- `OperationFailed`

Export audit properties:

- operation id
- correlation id
- requested by
- template id
- schema version
- output format
- output file name, if safe
- sheet names
- section/region names
- source names
- row counts per source
- column counts per source
- whether workbook encryption was requested
- whether sheet/workbook protection was requested
- validation warnings/errors
- duration
- final file size

Import audit properties:

- operation id
- source label
- file name, if safe
- source type, such as path, stream, bytes, upload, URI, or batch
- sheet/table/range imported
- row count
- column count
- included/excluded columns
- derived columns
- conversion errors
- validation errors
- whether an open password was used
- duration

Sensitive values that must not be logged by default:

- passwords
- secrets
- connection strings
- source row values
- full request payloads
- full schema when it contains secret references or sensitive metadata
- generated workbook bytes
- personally identifiable values unless explicitly allowed by host policy

The WebAPI should add correlation/request metadata and pass it into the library:

```csharp
ExcelReport.FromSchema(schema)
    .WithSource("invoiceRows", rows)
    .WithAuditContext(new DocWeaveOperationContext(
        OperationId: operationId,
        OperationType: "Export",
        TemplateId: templateId,
        SchemaVersion: schema.SchemaVersion,
        RequestedBy: userName,
        CorrelationId: correlationId,
        StartedAt: DateTimeOffset.UtcNow))
    .ToStream(output)
    .Save();
```

### Cancellation

Long-running imports, exports, template operations, and delivery operations should accept `CancellationToken`.

Cancellation points should exist before or during:

- source loading
- schema validation
- data binding
- layout resolution
- row enumeration
- sheet rendering
- pivot/chart generation
- encryption
- stream/file writing
- external delivery such as email or Teams

Cancellation should produce a controlled failure result and audit event rather than a corrupt partial output.

### Progress Updates

Progress should be reported as structured events, not direct UI messages.

Common progress stages:

- `Queued`
- `LoadingSources`
- `ValidatingSchema`
- `TranslatingSchema`
- `BindingSources`
- `ResolvingLayout`
- `RenderingSheet`
- `RenderingRegion`
- `WritingWorkbook`
- `EncryptingWorkbook`
- `DeliveringOutput`
- `Completed`
- `Failed`
- `Cancelled`

The WebAPI can expose progress through one or more host-level mechanisms:

- polling endpoint, such as `GET /operations/{operationId}`
- SignalR
- server-sent events
- message queue/event bus
- database status table
- webhook callback
- audit sink only, for background/batch operations

The core library should only emit progress events. The host decides where those events go.

## Output Delivery Interfaces

Renderers should produce files/streams. Delivery should be a separate adapter layer so exporting to email, Teams, storage, or an HTTP response does not complicate Excel rendering.

```csharp
public interface IExportDeliveryTarget
{
    string Kind { get; }
    Task DeliverAsync(ExportDeliveryRequest request, CancellationToken cancellationToken = default);
}

public sealed record ExportDeliveryRequest(
    string OperationId,
    string FileName,
    string ContentType,
    Stream Content,
    IReadOnlyDictionary<string, object?> Metadata);
```

Potential delivery targets:

- HTTP response stream.
- Local or network file path.
- Blob/object storage.
- Database/document store.
- Email attachment.
- Email link to stored file.
- Microsoft Teams chat/channel attachment or link.
- SharePoint/OneDrive document library.
- Webhook callback.
- Message queue with file reference.

Delivery adapters should live outside the core renderer packages, for example:

- `DocWeave.Delivery.Email`
- `DocWeave.Delivery.Teams`
- `DocWeave.Delivery.Storage`
- `DocWeave.WebApi`

Delivery audit should record destination type and status, but not message body secrets, file bytes, or sensitive row data.

## Template Interfaces

```csharp
public interface ITemplateWorkbookService
{
    WorkbookDocument OpenTemplate(Stream template);
    void CopySheet(string sourceSheet, string targetSheet);
    TemplateRange GetNamedRange(string name);
}

public interface ITemplateBindingEngine
{
    void BindValues(TemplateBindingSpec binding, BindingContext context);
    void BindTable(TemplateBindingSpec binding, IRowSource source, BindingContext context);
    void BindRepeatingBlock(RepeatingBlockSpec block, BindingContext context);
}
```

Template support should cover:

- Copy source sheets.
- Populate named ranges.
- Populate Excel tables.
- Copy template ranges for repeated blocks.
- Preserve template formatting.
- Replace placeholders.
- Bind multiple data sources into one sheet.
- Bind one data source into multiple regions.

## Word And PowerPoint Future Interfaces

These should be separate public entry points but share sources, binding, styles, templates, charts, and validation concepts.

```csharp
public interface IWordReportBuilder
{
    IWordReportBuilder UseTemplate(string path);
    IWordReportBuilder ToFile(string path);
    IWordReportBuilder Value(string placeholder, object? value);
    IWordReportBuilder Table<T>(string placeholder, IEnumerable<T> rows);
    IWordReportBuilder Repeat<T>(string sectionName, IEnumerable<T> items);
    void Save();
}

public interface IPowerPointDeckBuilder
{
    IPowerPointDeckBuilder UseTemplate(string path);
    IPowerPointDeckBuilder ToFile(string path);
    IPowerPointDeckBuilder Slide(string layoutName, Action<ISlideBuilder> configure);
    IPowerPointDeckBuilder RepeatSlide<T>(string layoutName, IEnumerable<T> items);
    void Save();
}
```

## Implementation Phases

### Phase 1: Core Excel Table Export

- `WorkbookSpec`, `SheetSpec`, `ExportRegionSpec`.
- File/stream/byte-array outputs.
- `IEnumerable<T>`, `DataTable`, `DataSet`, dictionaries.
- Multiple regions per sheet.
- Basic placement, headers, columns, styles, formats.
- Basic formulas and totals.
- CSV import/export.

### Phase 2: Layout Composition

- Groups with defaults and per-region overrides.
- Repeating blocks.
- Template ranges.
- Empty-source behavior.
- Named ranges and Excel tables.
- Region collision validation.

### Phase 3: Advanced Excel

- Conditional formatting.
- Pivot tables and pivot caches.
- Charts and pivot charts.
- Images, hyperlinks, comments, validation, protection.
- Sparklines.

### Phase 4: Declarative Schema

- JSON schema loading.
- Schema validation.
- Schema-to-spec compiler.
- `WithSource`, `WithDataSet`, `WithValue`, `WithAsset`.
- YAML/XML later if useful.

### Phase 5: Additional Formats

- TSV/custom-delimited.
- Fixed-width.
- JSON array and JSON Lines.
- XML records.
- HTML and Markdown tables.
- Parquet only if an external dependency is acceptable.

### Phase 6: Document And Presentation Renderers

- Word report generation.
- PowerPoint deck generation.
- Template binding.
- Repeated document sections/slides.
- Tables, charts, images, and theme/style tokens.

## Key Risks

- OpenXML pivot tables and slicers are complex and compatibility-sensitive.
- Chart rendering needs careful range/table reference generation.
- Template copying can easily break relationships if not isolated behind services.
- Region collision and auto-placement rules need deterministic validation.
- Auto-fit column calculation is approximate without Excel.
- Word/PowerPoint should not be allowed to distort the Excel-first model too early.

## Recommended First Interfaces To Implement

Start with these because they unlock most scenarios:

- `IExcelReportStart`
- `IExcelWorkbookBuilder`
- `IExcelSheetBuilder`
- `IExportRegionCollectionBuilder`
- `IExportRegionBuilder`
- `IDataSetExportBuilder`
- `IColumnBuilder`
- `IHeaderBuilder`
- `IStyleBuilder`
- `IRowSource`
- `IRowSourceFactory`
- `IWorkbookRenderer`
- `IRegionLayoutEngine`
- `ITableRenderer`
- `ISpecValidator<WorkbookSpec>`

Then add:

- `IRepeatingBlockBuilder<T>`
- `IPivotTableBuilder`
- `IChartBuilder`
- `ILayoutSchemaLoader`
- `ISchemaCompiler`
