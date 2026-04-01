using System.Collections.Generic;
using System.Linq;

namespace ExcelLibrary;

public class Row
{
    private readonly List<Cell> cells = new List<Cell>();

    public Row(int index)
    {
        Index = index;
        Hidden = false;
    }

    public Row(int index, bool hidden)
    {
        Index = index;
        Hidden = hidden;
    }

    public int Index { get; set; }
    public bool Hidden { get; set; }
    public Sheet Sheet { get; set; }

    public IEnumerable<Cell> Cells
    {
        get
        {
            if (this.Sheet.Workbook.Options.IncludeHidden)
            {
                return this.cells.OrderBy(c => c.Column.Index);
            }
            else
            {
                return this.cells.Where(c => c.Column.Hidden == false).OrderBy(c => c.Column.Index);
            }
        }
    }

    public void AddCell(Cell cell)
    {
        Cell match = (from c in this.cells
                      where c.Column.Index == cell.Column.Index
                      select c).SingleOrDefault();

        if (match == null)
        {
            cell.Row = this;
            this.cells.Add(cell);
        }
    }

    public Cell Cell(int index)
    {
        if (this.Sheet.Workbook.Options.IncludeHidden)
        {
            return this.cells.SingleOrDefault(c => c.Column.Index == index);
        }
        else
        {
            return this.cells.SingleOrDefault(c => c.Column.Index == index && c.Column.Hidden == false);
        }
    }
}
