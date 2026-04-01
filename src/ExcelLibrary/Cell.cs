using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class Cell
    {
        private Row row;
        private Column column;
        private string value;
        private NumberFormat format;

        public Cell(string value)
        {
            this.value = value;
        }

        public Row Row
        {
            get { return this.row; }
            set { this.row = value; }
        }

        public Column Column
        {
            get { return this.column; }
            set { this.column = value; }
        }

        public NumberFormat Format
        {
            get { return this.format; }
            set { this.format = value; }
        }

        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
    }
}
