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
            get { return this.path.ToLower(); }
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
            if (this.workbook.Options.IncludeHidden)
            {
                return this.rows.SingleOrDefault(r => r.Index == index);
            }
            else
            {
                return this.rows.SingleOrDefault(r => r.Index == index && r.Hidden == false);
            }
        }

        public Column Column(int index)
        {
            if (this.workbook.Options.IncludeHidden)
            {
                return this.columns.SingleOrDefault(c => c.Index == index);
            }
            else
            {
                return this.columns.SingleOrDefault(c => c.Index == index && c.Hidden == false);
            }
        }

        public Cell Cell(int rowIndex, int columnIndex)
        {
            IEnumerable<Cell> cells = this.rows.SelectMany(r => r.Cells);
            Cell cell = FindCell(cells, rowIndex, columnIndex);

            return cell;
        }

        public Cell Cell(string name)
        {
            Match match = Regex.Match(name, @"([A-Z]+)(\d+)");
            string letters = match.Groups[1].Value;
            string numbers = match.Groups[2].Value;
            int columnIndex = GetColumnIndex(letters);
            int rowIndex = Convert.ToInt16(numbers);

            IEnumerable<Cell> cells = this.rows.SelectMany(r => r.Cells);
            Cell cell = FindCell(cells, rowIndex, columnIndex);

            return cell;
        }

        private Cell FindCell(IEnumerable<Cell> cells, int rowIndex, int columnIndex)
        {
            if (this.workbook.Options.IncludeHidden)
            {
                Cell cell = (from c in cells
                             where c.Row.Index == rowIndex &&
                                   c.Column.Index == columnIndex
                             select c).SingleOrDefault();

                return cell;
            }
            else
            {
                Cell cell = (from c in cells
                             where c.Row.Index == rowIndex &&
                                   c.Row.Hidden == false &&
                                   c.Column.Index == columnIndex &&
                                   c.Column.Hidden == false
                             select c).SingleOrDefault();

                return cell;
            }
        }

        public IEnumerable<Cell> Cells
        {
            get
            {
                if (this.workbook.Options.IncludeHidden)
                {
                    return this.rows.SelectMany(r => r.Cells);
                }
                else
                {
                    return this.rows.Where(r => r.Hidden == false).SelectMany(r => r.Cells);
                }
            }
        }

        public IEnumerable<Row> Rows
        {
            get
            {
                if (this.workbook.Options.IncludeHidden)
                {
                    return this.rows.OrderBy(r => r.Index);
                }
                else
                {
                    return this.rows.Where(r => r.Hidden == false).OrderBy(r => r.Index);
                }
            }
        }

        public IEnumerable<Column> Columns
        {
            get
            {
                if (this.workbook.Options.IncludeHidden)
                {
                    return this.columns.OrderBy(c => c.Index);
                }
                else
                {
                    return this.columns.Where(c => c.Hidden == false).OrderBy(c => c.Index);
                }
            }
        }

        public void Open()
        {
            using (ZipArchive archive = ZipFile.OpenRead(this.workbook.File))
            {
                ZipArchiveEntry entry = archive.Entries.FirstOrDefault(e => e.FullName == this.Path.ToLower());
                this.Load(entry);
            }
        }

        private void Load(ZipArchiveEntry entry)
        {
            XDocument document = XDocument.Load(entry.Open());
            XElement root = document.Root;
            XNamespace ns = NS_MAIN;

            // Find hidden columns
            List<int> hiddenColumns = GetHiddenColumns(root, ns);

            // Loop throgh rows
            XElement sheetData = root.Element(ns + "sheetData");
            foreach (XElement eRow in sheetData.Elements(ns + "row"))
            {
                // Skip empty rows
                if (eRow.Descendants(ns + "v").Count() == 0)
                    continue;

                // Set row properties
                XAttribute attr1 = eRow.Attribute("r");
                XAttribute attr2 = eRow.Attribute("hidden");
                int index = Convert.ToInt16(attr1.Value);
                bool hidden = (attr2 != null && attr2.Value == "1") ? true : false;
                Row row = new Row(index, hidden);

                // Loop through cells on row
                foreach (XElement eCell in eRow.Elements(ns + "c"))
                {
                    // Skip empty cells
                    XElement xValue = eCell.Element(ns + "v");
                    if (xValue == null)
                        continue;

                    // Get cell position
                    string position = eCell.Attribute("r").Value;

                    // Get if the cell has text
                    string hasText_str = "";
                    bool hasText = false;

                    if(eCell.Attribute("t") != null)
                    { 
                        hasText_str = eCell.Attribute("t").Value;
                        if (hasText_str == "s")
                            hasText = true;
                    }
                    
                    // Get cell type
                    string type = "";
                    if(eCell.Attribute("s") != null)
                        type = eCell.Attribute("s").Value;

                    Match match = Regex.Match(position, @"([A-Z]+)(\d+)");
                    string letters = match.Groups[1].Value;
                    string numbers = match.Groups[2].Value;
                    int columnIndex = GetColumnIndex(letters);
                    int rowIndex = Convert.ToInt16(numbers);

                    // Get cell content
                    int number;
                    string sharedString = string.Empty;

                    // only try to load from SharedStrings if cell contains a string
                    if(hasText)
                    {
                        number = Convert.ToInt32(xValue.Value);
                        this.workbook.SharedStrings.TryGetValue(number, out sharedString);
                    }
                    
                    else
                    {
                        sharedString = xValue.Value;
                    }


                    // Make column
                    Column column = GetColumn(columnIndex);
                    column.Hidden = (hiddenColumns.Contains(columnIndex)) ? true : false;
                    
                    // Make cell
                    Cell cell = new Cell(sharedString);
                    
                    cell.Column = column;
                    cell.Row = row;
                    cell.Type = ConvertStringToType(type);
                    cell.HasText = hasText;

                    // Add cell to row and column
                    row.AddCell(cell);
                    column.AddCell(cell);

                    // Add rows and column to sheet
                    this.AddRow(row);
                    this.AddColumn(column);

                    /* We add rows and columns multiple times here and let
                     * the add methods filter out existing ones. */
                }
            }
        }

        private ExcelLibrary.Column GetColumn(int columnIndex)
        {
            // Try to find an existing column with the same index
            Column column = this.columns.SingleOrDefault(c => c.Index == columnIndex);
            
            if (column != null)
            {
                return column;
            }
            else
            {
                return new Column(columnIndex);
            }
        }

        private List<int> GetHiddenColumns(XElement root, XNamespace ns)
        {
            List<int> hiddenColumns = new List<int>();
            XElement eCols = root.Element(ns + "cols");

            if (eCols != null)
            {
                foreach (XElement eCol in eCols.Elements(ns + "col"))
                {
                    XAttribute aMin = eCol.Attribute("min");
                    XAttribute aMax = eCol.Attribute("max");
                    XAttribute aHidden = eCol.Attribute("hidden");

                    int min = (aMin != null) ? Convert.ToInt16(aMin.Value) : 0;
                    int max = (aMax != null) ? Convert.ToInt16(aMax.Value) : 0;
                    bool hidden = (aHidden != null && aHidden.Value == "1") ? true : false;
                    
                    if (hidden == false)
                        continue;

                    for (int i = min; i <= max; i++)
                    {
                        hiddenColumns.Add(i);
                    }
                }
            }

            return hiddenColumns;
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

        public string ConvertStringToType(string typeString)
        {
            switch (typeString)
            {
                case "1":
                    return "Number";
                case "2":
                    return "Currency";
                case "3":
                    return "Accounting";
                case "4":
                    return "Date";
                case "5":
                    return "Time";
                case "6":
                    return "Percentage";
                case "7":
                    return "Fraction";
                case "8":
                    return "Scientific";
                case "9":
                    return "Text";
                default:
                    return "General";
            }
        }
    }
}
