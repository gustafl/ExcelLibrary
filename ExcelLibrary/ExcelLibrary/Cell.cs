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
        private string type;
        public bool HasText { get; set; }

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

        public string Type
        {
            get { return this.type; }
            set { this.type = value; }
        }

        public string Value
        {
            get { return this.value; }
            set { this.value = value; }
        }
    }
}
