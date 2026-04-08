namespace ExcelLibrary;

/// <summary>
/// Represents a row in an Excel worksheet.
/// </summary>
/// <param name="index">The 1-based row index.</param>
/// <param name="hidden">Whether the row is hidden.</param>
public class Row(int index, bool hidden = false)
{
    private readonly Dictionary<int, Cell> cellsByColumnIndex = [];

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
    /// Gets the cells in this row.
    /// </summary>
    public IEnumerable<Cell> Cells =>
        Sheet.Workbook.Options.IncludeHidden
            ? cellsByColumnIndex.Values.OrderBy(c => c.Column.Index)
            : cellsByColumnIndex.Values.Where(c => !c.Column.Hidden).OrderBy(c => c.Column.Index);

    /// <summary>
    /// Adds a cell to this row.
    /// </summary>
    /// <param name="cell">The cell to add.</param>
    internal void AddCell(Cell cell) => cellsByColumnIndex.TryAdd(cell.Column.Index, cell);

    /// <summary>
    /// Gets a cell in this row by its column index.
    /// </summary>
    /// <param name="index">The 1-based column index.</param>
    /// <returns>The cell at the specified column, or <c>null</c> if not found or in a hidden column.</returns>
    public Cell? Cell(int index)
    {
        if (!cellsByColumnIndex.TryGetValue(index, out var cell))
            return null;
        return Sheet.Workbook.Options.IncludeHidden || !cell.Column.Hidden ? cell : null;
    }
}
