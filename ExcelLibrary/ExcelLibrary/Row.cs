using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class Row
    {
        private int index;
        private bool hidden;
        private Sheet sheet;
        private List<Cell> cells;

        public Row(int index)
        {
            this.index = index;
            this.hidden = false;
            this.cells = new List<Cell>();
        }

        public Row(int index, bool hidden)
        {
            this.index = index;
            this.hidden = hidden;
            this.cells = new List<Cell>();
        }

        public int Index
        {
            get { return this.index; }
            set { this.index = value; }
        }

        public bool Hidden
        {
            get { return this.hidden; }
            set { this.hidden = value; }
        }

        public Sheet Sheet
        {
            get { return this.sheet; }
            set { this.sheet = value; }
        }

        public IEnumerable<Cell> Cells
        {
            get { return this.cells.OrderBy(c => c.Column.Index); }
        }

        public void AddCell(Cell cell)
        {
            Cell match = (from c in this.cells
                          where c.Row.Index == cell.Row.Index &&
                                c.Column.Index == cell.Column.Index
                          select c).SingleOrDefault();

            if (match == null)
            {
                cell.Row = this;
                this.cells.Add(cell);
            }
        }

        public Cell Cell(int index)
        {
            Cell cell = this.cells.Where(c => c.Column.Index == index).SingleOrDefault();
            return cell;
        }

        public Cell Cell(string name)
        {
            throw new NotImplementedException();
        }
    }
}
