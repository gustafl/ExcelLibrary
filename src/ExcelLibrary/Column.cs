using System.Collections.Generic;
using System.Linq;

namespace ExcelLibrary;

public class Column
{
    private readonly List<Cell> cells = new List<Cell>();

    public Column(int index)
    {
        Index = index;
        Hidden = false;
    }

    public Column(int index, bool hidden)
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
                return this.cells.OrderBy(c => c.Row.Index);
            }
            else
            {
                return this.cells.Where(c => c.Row.Hidden == false).OrderBy(c => c.Row.Index);
            }
        }
    }

    public void AddCell(Cell cell)
    {
        Cell match = (from c in this.cells
                      where c.Row.Index == cell.Row.Index
                      select c).SingleOrDefault();

        if (match == null)
        {
            cell.Column = this;
            this.cells.Add(cell);
        }
    }

    public Cell Cell(int index)
    {
        if (this.Sheet.Workbook.Options.IncludeHidden)
        {
            return this.cells.SingleOrDefault(c => c.Row.Index == index);
        }
        else
        {
            return this.cells.SingleOrDefault(c => c.Row.Index == index && c.Row.Hidden == false);
        }
    }
}
