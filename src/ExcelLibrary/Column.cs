using System.Collections.Generic;
using System.Linq;

namespace ExcelLibrary;

public class Column(int index, bool hidden = false)
{
    private readonly List<Cell> cells = [];

    public int Index { get; set; } = index;
    public bool Hidden { get; set; } = hidden;
    public required Sheet Sheet { get; set; }

    public IEnumerable<Cell> Cells =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.OrderBy(c => c.Row.Index)
            : cells.Where(c => !c.Row.Hidden).OrderBy(c => c.Row.Index);

    public void AddCell(Cell cell)
    {
        if (cells.SingleOrDefault(c => c.Row.Index == cell.Row.Index) is null)
        {
            cell.Column = this;
            cells.Add(cell);
        }
    }

    public Cell? Cell(int index) =>
        Sheet.Workbook.Options.IncludeHidden
            ? cells.SingleOrDefault(c => c.Row.Index == index)
            : cells.SingleOrDefault(c => c.Row.Index == index && !c.Row.Hidden);
}
