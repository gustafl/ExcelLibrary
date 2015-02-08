using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelLibrary
{
    public class Sheet
    {
        private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

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

        public Row Row(int index)
        {
            Row row = this.rows.Where(r => r.Index == index).SingleOrDefault();
            return row;
        }

        public Column Column(int index)
        {
            Column column = this.columns.Where(c => c.Index == index).SingleOrDefault();
            return column;
        }

        public Cell Cell(int row, int column)
        {
            throw new NotImplementedException();

            // TODO: Find out how to get all cells from all rows using LINQ
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

        public void Load(ZipArchiveEntry entry)
        {
            XDocument document = XDocument.Load(entry.Open());
            XElement root = document.Root;
            XNamespace ns = NS_MAIN;

            // Loop throgh rows
            XElement sheetData = root.Element(ns + "sheetData");
            foreach (XElement eRow in sheetData.Elements(ns + "row"))
            {
                // Set row properties
                XAttribute attr1 = eRow.Attribute("r");
                XAttribute attr2 = eRow.Attribute("hidden");
                int index = Convert.ToInt16(attr1.Value);
                bool hidden = (attr2 != null && attr2.Value == "1") ? true : false;
                Row row = new Row(index, hidden);

                // Loop through cells on row
                foreach (XElement eCell in eRow.Elements(ns + "c"))
                {
                    // Get cell position
                    string position = eCell.Attribute("r").Value;
                    Match match = Regex.Match(position, @"([A-Z]+)(\d+)");
                    string letters = match.Groups[1].Value;
                    string numbers = match.Groups[2].Value;
                    int columnIndex = GetColumnIndex(letters);
                    int rowIndex = Convert.ToInt16(numbers);

                    // Get cell value
                    XElement xValue = eCell.Element(ns + "v");
                    if (xValue == null)
                        continue;

                    /* If the cell has not value (no <v> element), there's nothing more to do here.
                     * We are only collecting cells with content. */

                    int number = Convert.ToInt16(xValue.Value);
                    string sharedString = string.Empty;
                    this.workbook.SharedStrings.TryGetValue(number, out sharedString);

                    // Make column object
                    Column column = new Column(columnIndex);

                    // Add cell to row
                    Cell cell = new Cell(sharedString);
                    cell.Row = row;
                    cell.Column = column;
                    row.AddCell(cell);
                }

                // Add row to sheet
                this.AddRow(row);
            }
        }

        private int GetColumnIndex(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
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
