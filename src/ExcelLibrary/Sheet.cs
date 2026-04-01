namespace ExcelLibrary;

public partial class Sheet(string name, string? id = null, bool hidden = false)
{
    private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private string? path;
    private readonly List<Row> rows = [];
    private readonly List<Column> columns = [];

    public string Name { get; set; } = name;
    public string? Id { get; set; } = id;

    public string Path
    {
        get => path?.ToLower() ?? string.Empty;
        set => path = value;
    }

    public bool Hidden { get; set; } = hidden;
    public required Workbook Workbook { get; set; }

    public Row? Row(int index) =>
        Workbook.Options.IncludeHidden
            ? rows.SingleOrDefault(r => r.Index == index)
            : rows.SingleOrDefault(r => r.Index == index && !r.Hidden);

    public Column? Column(int index) =>
        Workbook.Options.IncludeHidden
            ? columns.SingleOrDefault(c => c.Index == index)
            : columns.SingleOrDefault(c => c.Index == index && !c.Hidden);

    public Cell? Cell(int rowIndex, int columnIndex) =>
        FindCell(rows.SelectMany(r => r.Cells), rowIndex, columnIndex);

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

    public IEnumerable<Cell> Cells =>
        Workbook.Options.IncludeHidden
            ? rows.SelectMany(r => r.Cells)
            : rows.Where(r => !r.Hidden).SelectMany(r => r.Cells);

    public IEnumerable<Row> Rows =>
        Workbook.Options.IncludeHidden
            ? rows.OrderBy(r => r.Index)
            : rows.Where(r => !r.Hidden).OrderBy(r => r.Index);

    public IEnumerable<Column> Columns =>
        Workbook.Options.IncludeHidden
            ? columns.OrderBy(c => c.Index)
            : columns.Where(c => !c.Hidden).OrderBy(c => c.Index);

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
