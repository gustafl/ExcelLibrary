namespace ExcelLibrary;

public class Row(int index, bool hidden = false)
{
    private readonly List<Cell> cells = [];

    public int Index { get; set; } = index;
    public bool Hidden { get; set; } = hidden;
    public required Sheet Sheet { get; set; }

    public IEnumerable<Cell> Cells =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.OrderBy(c => c.Column.Index)
            : cells.Where(c => !c.Column.Hidden).OrderBy(c => c.Column.Index);

    public void AddCell(Cell cell)
    {
        if (cells.SingleOrDefault(c => c.Column.Index == cell.Column.Index) is null)
        {
            cell.Row = this;
            cells.Add(cell);
        }
    }

    public Cell? Cell(int index) =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.SingleOrDefault(c => c.Column.Index == index)
            : cells.SingleOrDefault(c => c.Column.Index == index && !c.Column.Hidden);
}
