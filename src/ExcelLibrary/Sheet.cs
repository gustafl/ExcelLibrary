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

    private string? path;
    private readonly List<Row> rows = [];
    private readonly List<Column> columns = [];

    /// <summary>
    /// Gets or sets the name of the sheet as displayed in Excel.
    /// </summary>
    public string Name { get; set; } = name;

    /// <summary>
    /// Gets or sets the relationship ID used internally by Excel.
    /// </summary>
    public string? Id { get; set; } = id;

    /// <summary>
    /// Gets or sets the path to the sheet's XML file within the workbook archive.
    /// </summary>
    public string Path
    {
        get => path?.ToLower() ?? string.Empty;
        set => path = value;
    }

    /// <summary>
    /// Gets or sets whether this sheet is hidden in Excel.
    /// </summary>
    public bool Hidden { get; set; } = hidden;

    /// <summary>
    /// Gets the parent workbook containing this sheet.
    /// </summary>
    public required Workbook Workbook { get; set; }

    /// <summary>
    /// Gets a row by its 1-based index.
    /// </summary>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>The row at the specified index, or <c>null</c> if not found or hidden (when <see cref="WorkbookOptions.IncludeHidden"/> is <c>false</c>).</returns>
    public Row? Row(int index) =>
        Workbook.Options.IncludeHidden
            ? rows.SingleOrDefault(r => r.Index == index)
            : rows.SingleOrDefault(r => r.Index == index && !r.Hidden);

    /// <summary>
    /// Gets a column by its 1-based index.
    /// </summary>
    /// <param name="index">The 1-based column index (1 = A, 2 = B, etc.).</param>
    /// <returns>The column at the specified index, or <c>null</c> if not found or hidden (when <see cref="WorkbookOptions.IncludeHidden"/> is <c>false</c>).</returns>
    public Column? Column(int index) =>
        Workbook.Options.IncludeHidden
            ? columns.SingleOrDefault(c => c.Index == index)
            : columns.SingleOrDefault(c => c.Index == index && !c.Hidden);

    /// <summary>
    /// Gets a cell by its row and column indices.
    /// </summary>
    /// <param name="rowIndex">The 1-based row index.</param>
    /// <param name="columnIndex">The 1-based column index.</param>
    /// <returns>The cell at the specified position, or <c>null</c> if not found or in a hidden row/column.</returns>
    public Cell? Cell(int rowIndex, int columnIndex) =>
        FindCell(rows.SelectMany(r => r.Cells), rowIndex, columnIndex);

    /// <summary>
    /// Gets a cell by its Excel-style address (e.g., "A1", "B2", "AA10").
    /// </summary>
    /// <param name="name">The cell address in Excel notation.</param>
    /// <returns>The cell at the specified address, or <c>null</c> if not found or in a hidden row/column.</returns>
    public Cell? Cell(string name)
    {
        var match = CellAddressRegex().Match(name);
        int columnIndex = GetColumnIndex(match.Groups[1].Value);
        int rowIndex = int.Parse(match.Groups[2].Value);
        return FindCell(rows.SelectMany(r => r.Cells), rowIndex, columnIndex);
    }

    private Cell? FindCell(IEnumerable<Cell> cells, int rowIndex, int columnIndex) =>
        Workbook.Options.IncludeHidden
            ? cells.SingleOrDefault(c => c.Row.Index == rowIndex && c.Column.Index == columnIndex)
            : cells.SingleOrDefault(c => c.Row.Index == rowIndex && !c.Row.Hidden &&
                                         c.Column.Index == columnIndex && !c.Column.Hidden);

    /// <summary>
    /// Gets all cells in the sheet. Hidden cells are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Cell> Cells =>
        Workbook.Options.IncludeHidden
            ? rows.SelectMany(r => r.Cells)
            : rows.Where(r => !r.Hidden).SelectMany(r => r.Cells);

    /// <summary>
    /// Gets all rows in the sheet. Hidden rows are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Row> Rows =>
        Workbook.Options.IncludeHidden
            ? rows.OrderBy(r => r.Index)
            : rows.Where(r => !r.Hidden).OrderBy(r => r.Index);

    /// <summary>
    /// Gets all columns in the sheet. Hidden columns are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Column> Columns =>
        Workbook.Options.IncludeHidden
            ? columns.OrderBy(c => c.Index)
            : columns.Where(c => !c.Hidden).OrderBy(c => c.Index);

    /// <summary>
    /// Loads the sheet's data from the workbook file. Call this method when <see cref="WorkbookOptions.LoadSheets"/> is <c>false</c>.
    /// </summary>
    public void Open()
    {
        using var archive = ZipFile.OpenRead(Workbook.File!);
        var entry = archive.Entries.FirstOrDefault(e => e.FullName == Path.ToLower());
        if (entry is null) return;

        var document = XDocument.Load(entry.Open());
        var root = document.Root;
        if (root is null) return;

        XNamespace ns = NS_MAIN;

        // Find hidden columns
        var hiddenColumns = GetHiddenColumns(root, ns);

        // Loop through rows
        var sheetData = root.Element(ns + "sheetData");
        if (sheetData is null) return;

        foreach (var eRow in sheetData.Elements(ns + "row"))
        {
            // Skip empty rows
            if (!eRow.Descendants(ns + "v").Any())
                continue;

            // Set row properties
            int index = int.Parse(eRow.Attribute("r")!.Value);
            bool hidden = eRow.Attribute("hidden") is { Value: "1" };
            var row = new Row(index, hidden) { Sheet = this };

            // Loop through cells on row
            foreach (var eCell in eRow.Elements(ns + "c"))
            {
                // Skip empty cells
                var xValue = eCell.Element(ns + "v");
                if (xValue is null)
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

                // Get shared string (if any)
                string? sharedString = null;
                if (IsSharedString(eCell))
                {
                    int number = int.Parse(xValue.Value);
                    Workbook.SharedStrings.TryGetValue(number, out sharedString);
                }

                // Make column
                var column = GetColumn(columnIndex);
                column.Hidden = hiddenColumns.Contains(columnIndex);

                // Make cell
                string cellValue = sharedString ?? xValue.Value;
                var cell = new Cell(cellValue) { Column = column, Row = row, Format = format };

                if (format is NumberFormat.Date)
                    cell.Value = Utilities.ConvertDate(cell.Value, Workbook.BaseYear);
                else if (format is NumberFormat.Time)
                    cell.Value = Utilities.ConvertTime(cell.Value);

                // Add cell to row and column
                row.AddCell(cell);
                column.AddCell(cell);

                // Add rows and column to sheet
                AddRow(row);
                AddColumn(column);
            }
        }
    }

    private bool IsSharedString(XElement cell) => cell.Attribute("t") is { Value: "s" };

    private Column GetColumn(int columnIndex) =>
        columns.SingleOrDefault(c => c.Index == columnIndex) ?? new Column(columnIndex) { Sheet = this };

    private List<int> GetHiddenColumns(XElement root, XNamespace ns)
    {
        List<int> hiddenColumns = [];
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

    private static int GetColumnIndex(string name)
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

    public void AddRow(Row row)
    {
        if (rows.SingleOrDefault(r => r.Index == row.Index) is null)
        {
            row.Sheet = this;
            rows.Add(row);
        }
    }

    public void AddColumn(Column column)
    {
        if (columns.SingleOrDefault(c => c.Index == column.Index) is null)
        {
            column.Sheet = this;
            columns.Add(column);
        }
    }
}
