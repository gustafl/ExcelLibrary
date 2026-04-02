namespace ExcelLibrary;

/// <summary>
/// Represents a row in an Excel worksheet.
/// </summary>
/// <param name="index">The 1-based row index.</param>
/// <param name="hidden">Whether the row is hidden.</param>
public class Row(int index, bool hidden = false)
{
    private readonly List<Cell> cells = [];

    /// <summary>
    /// Gets the 1-based row index.
    /// </summary>
    public int Index { get; } = index;

    /// <summary>
    /// Gets whether this row is hidden in Excel.
    /// </summary>
    public bool Hidden { get; } = hidden;

    /// <summary>
    /// Gets the parent sheet containing this row.
    /// </summary>
    public required Sheet Sheet { get; init; }

    /// <summary>
    /// Gets the cells in this row. Cells in hidden columns are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Cell> Cells =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.OrderBy(c => c.Column.Index)
            : cells.Where(c => !c.Column.Hidden).OrderBy(c => c.Column.Index);

    /// <summary>
    /// Adds a cell to this row.
    /// </summary>
    /// <param name="cell">The cell to add.</param>
    internal void AddCell(Cell cell)
    {
        if (cells.SingleOrDefault(c => c.Column.Index == cell.Column.Index) is null)
        {
            cells.Add(cell);
        }
    }

    /// <summary>
    /// Gets a cell in this row by its column index.
    /// </summary>
    /// <param name="index">The 1-based column index.</param>
    /// <returns>The cell at the specified column, or <c>null</c> if not found or in a hidden column.</returns>
    public Cell? Cell(int index) =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.SingleOrDefault(c => c.Column.Index == index)
            : cells.SingleOrDefault(c => c.Column.Index == index && !c.Column.Hidden);
}
