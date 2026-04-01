namespace ExcelLibrary;

/// <summary>
/// Represents a column in an Excel worksheet.
/// </summary>
/// <param name="index">The 1-based column index (1 = A, 2 = B, etc.).</param>
/// <param name="hidden">Whether the column is hidden.</param>
public class Column(int index, bool hidden = false)
{
    private readonly List<Cell> cells = [];

    /// <summary>
    /// Gets or sets the 1-based column index (1 = A, 2 = B, etc.).
    /// </summary>
    public int Index { get; set; } = index;

    /// <summary>
    /// Gets or sets whether this column is hidden in Excel.
    /// </summary>
    public bool Hidden { get; set; } = hidden;

    /// <summary>
    /// Gets the parent sheet containing this column.
    /// </summary>
    public required Sheet Sheet { get; set; }

    /// <summary>
    /// Gets the cells in this column. Cells in hidden rows are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Cell> Cells =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.OrderBy(c => c.Row.Index)
            : cells.Where(c => !c.Row.Hidden).OrderBy(c => c.Row.Index);

    /// <summary>
    /// Adds a cell to this column. Internal use only.
    /// </summary>
    /// <param name="cell">The cell to add.</param>
    public void AddCell(Cell cell)
    {
        if (cells.SingleOrDefault(c => c.Row.Index == cell.Row.Index) is null)
        {
            cell.Column = this;
            cells.Add(cell);
        }
    }

    /// <summary>
    /// Gets a cell in this column by its row index.
    /// </summary>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>The cell at the specified row, or <c>null</c> if not found or in a hidden row.</returns>
    public Cell? Cell(int index) =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.SingleOrDefault(c => c.Row.Index == index)
            : cells.SingleOrDefault(c => c.Row.Index == index && !c.Row.Hidden);
}
