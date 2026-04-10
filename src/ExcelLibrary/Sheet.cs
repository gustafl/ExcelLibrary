namespace ExcelLibrary;

/// <summary>
/// Represents a worksheet within an Excel workbook.
/// </summary>
/// <param name="name">The name of the sheet.</param>
/// <param name="id">The relationship ID of the sheet.</param>
/// <param name="hidden">Whether the sheet is hidden.</param>
public partial class Sheet(string name, string? id = null, bool hidden = false)
{
    private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private readonly Dictionary<int, Row> rowsByIndex = [];
    private readonly Dictionary<int, Column> columnsByIndex = [];

    /// <summary>
    /// Gets the name of the sheet as displayed in Excel.
    /// </summary>
    public string Name { get; } = name;

    /// <summary>
    /// Gets the relationship ID used internally by Excel.
    /// </summary>
    public string? Id { get; } = id;

    /// <summary>
    /// Gets the path to the sheet's XML file within the workbook archive.
    /// </summary>
    public required string Path { get; init; }

    /// <summary>
    /// Gets whether this sheet is hidden in Excel.
    /// </summary>
    public bool Hidden { get; } = hidden;

    /// <summary>
    /// Gets the parent workbook containing this sheet.
    /// </summary>
    public required Workbook Workbook { get; init; }

    /// <summary>
    /// Gets whether this sheet's data has been loaded.
    /// </summary>
    internal bool IsLoaded { get; private set; }

    /// <summary>
    /// Gets a row by its 1-based index.
    /// </summary>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>The row at the specified index, or <c>null</c> if not found or hidden (when <see cref="WorkbookOptions.IncludeHidden"/> is <c>false</c>).</returns>
    public Row? Row(int index)
    {
        if (!rowsByIndex.TryGetValue(index, out var row))
            return null;
        return Workbook.Options.IncludeHidden || !row.Hidden ? row : null;
    }

    /// <summary>
    /// Gets a column by its 1-based index.
    /// </summary>
    /// <param name="index">The 1-based column index (1 = A, 2 = B, etc.).</param>
    /// <returns>The column at the specified index, or <c>null</c> if not found or hidden (when <see cref="WorkbookOptions.IncludeHidden"/> is <c>false</c>).</returns>
    public Column? Column(int index)
    {
        if (!columnsByIndex.TryGetValue(index, out var column))
            return null;
        return Workbook.Options.IncludeHidden || !column.Hidden ? column : null;
    }

    /// <summary>
    /// Gets a cell by its row and column indices.
    /// </summary>
    /// <param name="rowIndex">The 1-based row index.</param>
    /// <param name="columnIndex">The 1-based column index.</param>
    /// <returns>The cell at the specified position, or <c>null</c> if not found or in a hidden row/column.</returns>
    public Cell? Cell(int rowIndex, int columnIndex)
    {
        var row = Row(rowIndex);
        return row?.Cell(columnIndex);
    }

    /// <summary>
    /// Gets a cell by its Excel-style address (e.g., "A1", "B2", "AA10").
    /// </summary>
    /// <param name="name">The cell address in Excel notation.</param>
    /// <returns>The cell at the specified address, or <c>null</c> if not found or in a hidden row/column.</returns>
    public Cell? Cell(string name)
    {
        var match = CellAddressRegex().Match(name);
        int columnIndex = GetColumnIndex(match.Groups[1].ValueSpan);
        int rowIndex = int.Parse(match.Groups[2].ValueSpan);
        return Cell(rowIndex, columnIndex);
    }

    /// <summary>
    /// Gets all cells in the sheet. Hidden cells are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Cell> Cells =>
        Workbook.Options.IncludeHidden
            ? rowsByIndex.Values.SelectMany(r => r.Cells)
            : rowsByIndex.Values.Where(r => !r.Hidden).SelectMany(r => r.Cells);

    /// <summary>
    /// Gets all rows in the sheet. Hidden rows are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Row> Rows =>
        Workbook.Options.IncludeHidden
            ? rowsByIndex.Values.OrderBy(r => r.Index)
            : rowsByIndex.Values.Where(r => !r.Hidden).OrderBy(r => r.Index);

    /// <summary>
    /// Gets all columns in the sheet. Hidden columns are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Column> Columns =>
        Workbook.Options.IncludeHidden
            ? columnsByIndex.Values.OrderBy(c => c.Index)
            : columnsByIndex.Values.Where(c => !c.Hidden).OrderBy(c => c.Index);

    /// <summary>
    /// Loads the sheet's data from the workbook file. Call this method when <see cref="WorkbookOptions.LoadSheets"/> is <c>false</c>.
    /// </summary>
    public void Open()
    {
        if (IsLoaded) return;

        using var stream = Workbook.OpenArchiveEntry(Path);
        if (stream is null) return;

        var document = XDocument.Load(stream);
        var root = document.Root;
        if (root is null) return;

        XNamespace ns = NS_MAIN;

        // Find hidden columns
        var hiddenColumns = GetHiddenColumns(root, ns);

        // Loop through rows
        var sheetData = root.Element(ns + "sheetData");
        if (sheetData is null)
        {
            IsLoaded = true;
            Workbook.NotifySheetLoaded();
            return;
        }

        foreach (var eRow in sheetData.Elements(ns + "row"))
        {
            // Skip empty rows (no <v> values and no inline strings)
            bool hasValues = eRow.Descendants(ns + "v").Any();
            bool hasInlineStrings = eRow.Elements(ns + "c").Any(c => c.Attribute("t") is { Value: "inlineStr" });
            if (!hasValues && !hasInlineStrings)
                continue;

            // Set row properties
            int index = int.Parse(eRow.Attribute("r")!.Value);
            bool hidden = eRow.Attribute("hidden") is { Value: "1" };
            var row = new Row(index, hidden) { Sheet = this };

            // Loop through cells on row
            foreach (var eCell in eRow.Elements(ns + "c"))
            {
                // Skip empty cells (no <v> value and no inline string)
                var xValue = eCell.Element(ns + "v");
                bool hasInlineString = IsInlineString(eCell);
                if (xValue is null && !hasInlineString)
                    continue;

                // Get cell position
                var match = CellAddressRegex().Match(eCell.Attribute("r")!.Value);
                int columnIndex = GetColumnIndex(match.Groups[1].Value);
                int rowIndex = int.Parse(match.Groups[2].Value);

                // Get cell style
                var format = NumberFormat.General;
                if (eCell.Attribute("s") is { } s)
                {
                    int styleIndex = int.Parse(s.Value);
                    format = (NumberFormat)Workbook.NumberFormats[styleIndex];
                }

                // Resolve string value: shared string, inline string, or direct value
                string? resolvedString = null;
                if (IsSharedString(eCell))
                {
                    int number = int.Parse(xValue!.Value);
                    Workbook.SharedStrings.TryGetValue(number, out resolvedString);
                }
                else if (hasInlineString)
                {
                    resolvedString = GetInlineStringValue(eCell);
                }

                // Make column
                bool isHiddenColumn = hiddenColumns.Contains(columnIndex);
                var column = GetOrCreateColumn(columnIndex, isHiddenColumn);

                // Compute cell value (with format conversion)
                string rawValue = resolvedString ?? xValue?.Value ?? string.Empty;
                string cellValue = format switch
                {
                    NumberFormat.Date => Utilities.ConvertDate(rawValue, Workbook.BaseYear),
                    NumberFormat.Time => Utilities.ConvertTime(rawValue),
                    _ => rawValue
                };

                // Make cell
                var cell = new Cell(cellValue) { Column = column, Row = row, Format = format };

                // Add cell to row and column
                row.AddCell(cell);
                column.AddCell(cell);

                // Add rows and column to sheet
                AddRow(row);
                AddColumn(column);
            }
        }

        IsLoaded = true;
        Workbook.NotifySheetLoaded();
    }

    private bool IsSharedString(XElement cell) => cell.Attribute("t") is { Value: "s" };

    private bool IsInlineString(XElement cell) => cell.Attribute("t") is { Value: "inlineStr" };

    private static string? GetInlineStringValue(XElement cell)
    {
        // Inline string structure: <c t="inlineStr"><is><t>value</t></is></c>
        var ns = cell.Name.Namespace;
        return cell.Element(ns + "is")?.Element(ns + "t")?.Value;
    }

    private Column GetOrCreateColumn(int columnIndex, bool hidden)
    {
        if (columnsByIndex.TryGetValue(columnIndex, out var existing))
            return existing;
        return new Column(columnIndex, hidden) { Sheet = this };
    }

    private HashSet<int> GetHiddenColumns(XElement root, XNamespace ns)
    {
        HashSet<int> hiddenColumns = [];
        var eCols = root.Element(ns + "cols");

        if (eCols is null)
            return hiddenColumns;

        foreach (var eCol in eCols.Elements(ns + "col"))
        {
            int min = eCol.Attribute("min") is { } aMin ? int.Parse(aMin.Value) : 0;
            int max = eCol.Attribute("max") is { } aMax ? int.Parse(aMax.Value) : 0;
            bool hidden = eCol.Attribute("hidden") is { Value: "1" };

            if (!hidden)
                continue;

            for (int i = min; i <= max; i++)
                hiddenColumns.Add(i);
        }

        return hiddenColumns;
    }

    [GeneratedRegex(@"([A-Z]+)(\d+)")]
    private static partial Regex CellAddressRegex();

    private static int GetColumnIndex(ReadOnlySpan<char> name)
    {
        int number = 0;
        int pow = 1;
        for (int i = name.Length - 1; i >= 0; i--)
        {
            number += (name[i] - 'A' + 1) * pow;
            pow *= 26;
        }

        return number;
    }

    internal void AddRow(Row row) => rowsByIndex.TryAdd(row.Index, row);

    internal void AddColumn(Column column) => columnsByIndex.TryAdd(column.Index, column);
}



