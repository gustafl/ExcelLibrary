using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public enum SheetVisibility { Visible, Hidden };

    public class Sheet
    {
        private string name;
        private string id;
        private string path;
        private bool hidden;
        private Workbook workbook;
        private List<Row> rows;
        private List<Column> columns;

        public Sheet(string name)
        {
            this.name = name;
            this.rows = new List<Row>();
            this.columns = new List<Column>();
        }

        public Sheet(string name, string id, bool hidden)
        {
            this.name = name;
            this.id = id;
            this.hidden = hidden;
            this.rows = new List<Row>();
            this.columns = new List<Column>();
        }

        public string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        public string Id
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public string Path
        {
            get { return this.path; }
            set { this.path = value; }
        }

        public bool Hidden
        {
            get { return this.hidden; }
            set { this.hidden = value; }
        }

        public Workbook Workbook
        {
            get { return this.workbook; }
            set { this.workbook = value; }
        }

        public IEnumerable<Row> Rows
        {
            get
            {
                List<Row> rowsToReturn = new List<Row>();

                IEnumerable<Row> visibleRows = this.rows.Where(r => r.Hidden == false);
                rowsToReturn.AddRange(visibleRows);

                if (this.workbook.Options.IncludeHidden)
                {
                    IEnumerable<Row> hiddenRows = this.rows.Where(r => r.Hidden == true);
                    rowsToReturn.AddRange(hiddenRows);
                }

                return rowsToReturn;
            }
        }

        public IEnumerable<Column> Columns
        {
            get
            {
                throw new NotImplementedException();    // TODO
            }
        }

        public void AddRow(Row row)
        {
            Row match = (from r in this.rows
                         where r.Index == row.Index
                         select r).SingleOrDefault();

            if (match == null)
            {
                row.Sheet = this;
                this.rows.Add(row);
            }
        }

        public void AddColumn(Column column)
        {
            Column match = (from c in this.columns
                            where c.Index == column.Index
                            select c).SingleOrDefault();

            if (match == null)
            {
                column.Sheet = this;
                this.columns.Add(column);
            }
        }
    }
}
