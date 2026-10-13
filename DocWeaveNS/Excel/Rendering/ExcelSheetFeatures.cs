using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace DocWeave.Excel;

/// <summary>
/// Worksheet-level features that are not cell content: views, protection, comments, data validation and native Excel tables.
/// </summary>
internal static class ExcelSheetFeatures
{
    internal static SheetViews CreateSheetViews(ExcelFreezePaneSpec freezePane)
    {
        // The scrolling area starts at the first cell below the frozen rows and to the right of the frozen columns.
        var topLeft = new ExcelCellAddress(freezePane.FrozenRows + 1, freezePane.FrozenColumns + 1).ToReference();
        var pane = new Pane
        {
            State = PaneStateValues.Frozen,
            TopLeftCell = topLeft
        };

        if (freezePane.FrozenColumns > 0)
        {
            pane.HorizontalSplit = freezePane.FrozenColumns;
        }

        if (freezePane.FrozenRows > 0)
        {
            pane.VerticalSplit = freezePane.FrozenRows;
        }

        // The active pane is where the cursor sits: the bottom-right pane when both rows and columns are frozen,
        // otherwise the only pane that scrolls (bottom-left for frozen rows, top-right for frozen columns).
        pane.ActivePane = freezePane.FrozenRows > 0 && freezePane.FrozenColumns > 0
            ? PaneValues.BottomRight
            : freezePane.FrozenRows > 0
                ? PaneValues.BottomLeft
                : PaneValues.TopRight;

        return new SheetViews(new SheetView(
            pane,
            new Selection { Pane = pane.ActivePane, ActiveCell = topLeft, SequenceOfReferences = new ListValue<StringValue> { InnerText = topLeft } })
        {
            WorkbookViewId = 0U
        });
    }

    internal static WorkbookProtection CreateWorkbookProtection(ExcelWorkbookProtectionSpec protection)
    {
        var workbookProtection = new WorkbookProtection
        {
            LockStructure = protection.LockStructure
        };

        if (!string.IsNullOrEmpty(protection.Password))
        {
            workbookProtection.WorkbookPassword = HashProtectionPassword(protection.Password);
        }

        return workbookProtection;
    }

    internal static SheetProtection CreateSheetProtection(ExcelSheetProtectionSpec protection)
    {
        var sheetProtection = new SheetProtection
        {
            // The SelectLockedCells, Sort and AutoFilter attributes mean "blocked" when true, so the caller's Allow* flags are inverted.
            Sheet = true,
            Objects = true,
            Scenarios = true,
            SelectLockedCells = !protection.AllowSelectLockedCells,
            Sort = !protection.AllowSort,
            AutoFilter = !protection.AllowAutoFilter
        };

        if (!string.IsNullOrEmpty(protection.Password))
        {
            sheetProtection.Password = HashProtectionPassword(protection.Password);
        }

        return sheetProtection;
    }

    private static string HashProtectionPassword(string password)
    {
        // Excel's legacy password hash, the one it reads from files: a 15-bit value built by rotating and XOR-ing each
        // character from last to first, then the length and a fixed constant. It is easy to break, so it only stops
        // accidental edits. It is not a way to keep anything secret.
        var hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= password[i];
        }

        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;
        return hash.ToString("X4", CultureInfo.InvariantCulture);
    }

    internal static void AddComments(WorksheetPart worksheetPart, ExcelSheetSpec sheetSpec)
    {
        var commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
        // Each comment names its author by position in a shared list of authors, so build that list first.
        var authors = new Authors();
        var authorIndexes = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);

        foreach (var author in sheetSpec.Comments.Select(comment => comment.Author).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            authorIndexes[author] = (uint)authorIndexes.Count;
            authors.Append(new Author(author));
        }

        var commentList = new CommentList();
        foreach (var comment in sheetSpec.Comments)
        {
            var spreadsheetComment = new Comment
            {
                Reference = comment.Address.ToReference(),
                AuthorId = authorIndexes[comment.Author]
            };
            spreadsheetComment.Append(new CommentText(new Run(new Text(comment.Text))));
            commentList.Append(spreadsheetComment);
        }

        commentsPart.Comments = new Comments(authors, commentList);
        commentsPart.Comments.Save();
    }

    internal static DataValidations CreateDataValidations(ExcelSheetSpec sheetSpec)
    {
        var dataValidations = new DataValidations { Count = (uint)sheetSpec.DataValidations.Count };
        foreach (var validation in sheetSpec.DataValidations)
        {
            var dataValidation = new DataValidation
            {
                Type = MapDataValidationType(validation.Kind),
                // List and custom rules are not comparisons, so they have no operator.
                Operator = validation.Kind == ExcelDataValidationKind.List || validation.Kind == ExcelDataValidationKind.Custom
                    ? null
                    : MapDataValidationOperator(validation.Operator),
                SequenceOfReferences = new ListValue<StringValue> { InnerText = validation.Range },
                AllowBlank = validation.AllowBlank,
                // The input tip and error box are switched on only when the caller supplied text for them.
                ShowInputMessage = validation.InputMessage is not null || validation.InputTitle is not null,
                ShowErrorMessage = validation.ErrorMessage is not null || validation.ErrorTitle is not null,
                PromptTitle = validation.InputTitle,
                Prompt = validation.InputMessage,
                ErrorTitle = validation.ErrorTitle,
                Error = validation.ErrorMessage
            };

            if (validation.Formula1 is not null)
            {
                dataValidation.Append(new Formula1(validation.Formula1));
            }

            if (validation.Formula2 is not null)
            {
                dataValidation.Append(new Formula2(validation.Formula2));
            }

            dataValidations.Append(dataValidation);
        }

        return dataValidations;
    }

    internal static void AddNativeTables(WorksheetPart worksheetPart, ExcelSheetSpec sheetSpec)
    {
        var nativeTables = sheetSpec.Tables
            .Where(table => table.NativeTable is not null)
            .ToList();

        if (nativeTables.Count == 0)
        {
            return;
        }

        var tableParts = new TableParts { Count = (uint)nativeTables.Count };
        for (var i = 0; i < nativeTables.Count; i++)
        {
            var table = nativeTables[i];
            var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tablePart.Table = CreateNativeTableDefinition(table, (uint)(i + 1));
            tablePart.Table.Save();
            tableParts.Append(new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) });
        }

        worksheetPart.Worksheet.Append(tableParts);
    }

    private static DocumentFormat.OpenXml.Spreadsheet.Table CreateNativeTableDefinition(ExcelTableSpec table, uint tableId)
    {
        var spec = table.NativeTable!;
        var columns = table.Data.Columns.Concat(table.FormulaColumns.Select(column => column.ColumnName)).ToList();
        if (columns.Count == 0)
        {
            throw new InvalidOperationException($"Native Excel table '{spec.Name}' must contain at least one column.");
        }

        // Columns can be placed individually, so the table spans from the leftmost to the rightmost column actually used.
        var startColumn = columns
            .Select((column, index) => ExcelTableRenderer.ResolveColumnNumber(table, column, index))
            .Min();
        var endColumn = columns
            .Select((column, index) => ExcelTableRenderer.ResolveColumnNumber(table, column, index))
            .Max();
        // A native table is written with exactly one header row.
        var headerRows = 1;
        // An empty source still reserves one data row.
        var dataRowCount = Math.Max(table.Data.Rows.Count, 1);
        var dataEndRow = table.StartCell.Row + headerRows + dataRowCount - 1;
        var reference = $"{new ExcelCellAddress(table.StartCell.Row, startColumn).ToReference()}:{new ExcelCellAddress(dataEndRow, endColumn).ToReference()}";
        var filterReference = $"{new ExcelCellAddress(table.StartCell.Row, startColumn).ToReference()}:{new ExcelCellAddress(dataEndRow, endColumn).ToReference()}";

        var tableColumns = new TableColumns { Count = (uint)columns.Count };
        for (var i = 0; i < columns.Count; i++)
        {
            var sourceName = columns[i];
            // Use the caption shown in the header row (the alias, if any) as the table column name.
            var caption = table.HeaderAliases.TryGetValue(sourceName, out var alias) ? alias : sourceName;
            tableColumns.Append(new TableColumn
            {
                Id = (uint)(i + 1),
                Name = caption
            });
        }

        var nativeTable = new DocumentFormat.OpenXml.Spreadsheet.Table
        {
            Id = tableId,
            Name = ExcelNames.SanitizeTableName(spec.Name),
            DisplayName = ExcelNames.SanitizeTableName(spec.Name),
            Reference = reference,
            TotalsRowShown = false
        };

        if (spec.ShowFilterButtons)
        {
            nativeTable.Append(new AutoFilter { Reference = filterReference });
        }

        nativeTable.Append(tableColumns);
        nativeTable.Append(new TableStyleInfo
        {
            Name = spec.StyleName,
            ShowFirstColumn = false,
            ShowLastColumn = false,
            ShowRowStripes = spec.ShowRowStripes,
            ShowColumnStripes = false
        });

        return nativeTable;
    }

    private static DataValidationValues MapDataValidationType(ExcelDataValidationKind kind)
    {
        return kind switch
        {
            ExcelDataValidationKind.List => DataValidationValues.List,
            ExcelDataValidationKind.WholeNumber => DataValidationValues.Whole,
            ExcelDataValidationKind.Decimal => DataValidationValues.Decimal,
            ExcelDataValidationKind.Date => DataValidationValues.Date,
            ExcelDataValidationKind.TextLength => DataValidationValues.TextLength,
            ExcelDataValidationKind.Custom => DataValidationValues.Custom,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
        };
    }

    private static DataValidationOperatorValues MapDataValidationOperator(ExcelDataValidationOperator operatorKind)
    {
        return operatorKind switch
        {
            ExcelDataValidationOperator.Between => DataValidationOperatorValues.Between,
            ExcelDataValidationOperator.NotBetween => DataValidationOperatorValues.NotBetween,
            ExcelDataValidationOperator.Equal => DataValidationOperatorValues.Equal,
            ExcelDataValidationOperator.NotEqual => DataValidationOperatorValues.NotEqual,
            ExcelDataValidationOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
            ExcelDataValidationOperator.LessThan => DataValidationOperatorValues.LessThan,
            ExcelDataValidationOperator.GreaterThanOrEqual => DataValidationOperatorValues.GreaterThanOrEqual,
            ExcelDataValidationOperator.LessThanOrEqual => DataValidationOperatorValues.LessThanOrEqual,
            _ => throw new ArgumentOutOfRangeException(nameof(operatorKind), operatorKind, null)
        };
    }
}
