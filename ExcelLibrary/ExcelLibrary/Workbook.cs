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
    public class Workbook
    {
        private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NS_OR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NS_ORW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        private string file;
        private List<Sheet> sheets;
        private Dictionary<int, string> sharedStrings;
        private WorkbookOptions options;

        public void Open(string file)
        {
            this.file = file;
            this.sheets = new List<Sheet>();
            this.sharedStrings = new Dictionary<int, string>();
            this.options = new WorkbookOptions();
            Open();
        }

        public void Open(string file, WorkbookOptions options)
        {
            this.file = file;
            this.sheets = new List<Sheet>();
            this.sharedStrings = new Dictionary<int, string>();
            this.options = options;
            Open();
        }

        private void Open()
        {
            using (ZipArchive archive = ZipFile.OpenRead(file))
            {
                // Read "xl/workbook.xml" to get sheet names and ids
                ZipArchiveEntry sheetsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/workbook.xml");
                LoadWorkbookXml(sheetsEntry);

                // Read "xl/_rels/workbook.xml.rels" to get sheet paths
                ZipArchiveEntry sheetPathsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/_rels/workbook.xml.rels");
                LoadWorkbookXmlRels(sheetPathsEntry);

                // Read "xl/sharedStrings.xml" to get shared strings
                ZipArchiveEntry sharedStringsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/sharedStrings.xml");
                LoadSharedStrings(sharedStringsEntry);

                // Open each sheet found
                foreach (Sheet sheet in this.sheets)
                {
                    ZipArchiveEntry entry = archive.Entries.FirstOrDefault(e => e.FullName == sheet.Path);
                    LoadSheet(entry, sheet);
                }
            }
        }

        private void LoadWorkbookXml(ZipArchiveEntry entry)
        {
            XDocument document = XDocument.Load(entry.Open());
            XElement root = document.Root;
            XNamespace ns = NS_MAIN;
            XNamespace r = NS_OR;

            foreach (XElement element in root.Element(ns + "sheets").Elements())
            {
                XAttribute id = element.Attribute(r + "id");
                XAttribute name = element.Attribute("name");
                XAttribute state = element.Attribute("state");

                bool hidden = false;
                if (state != null && state.Value == "hidden")
                    hidden = true;

                Sheet sheet = new Sheet(name.Value, id.Value, hidden);
                sheet.Workbook = this;
                this.sheets.Add(sheet);
            }
        }

        private void LoadWorkbookXmlRels(ZipArchiveEntry entry)
        {
            XDocument document = XDocument.Load(entry.Open());
            XElement root = document.Root;
            XNamespace ns = NS_PR;

            foreach (XElement element in root.Elements(ns + "Relationship"))
            {
                XAttribute id = element.Attribute("Id");
                XAttribute type = element.Attribute("Type");
                XAttribute target = element.Attribute("Target");

                if (type.Value != NS_ORW)
                    continue;

                Sheet sheet = (from s in this.sheets
                               where s.Id == id.Value
                               select s).FirstOrDefault();

                // Get sheet file path
                Match match = Regex.Match(target.Value, @"worksheets/(.+)");
                string path = @"xl/worksheets/" + match.Groups[1].Value;
                sheet.Path = path;
            }
        }

        private void LoadSharedStrings(ZipArchiveEntry entry)
        {
            XDocument document = XDocument.Load(entry.Open());
            XElement root = document.Root;
            XNamespace ns = NS_MAIN;
            int count = 0;

            foreach (XElement si in root.Elements(ns + "si"))
            {
                XElement t = si.Element(ns + "t");
                this.sharedStrings.Add(count, t.Value);
                count++;
            }
        }

        private void LoadSheet(ZipArchiveEntry entry, Sheet sheet)
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
                    int number = Convert.ToInt16(xValue.Value);
                    string sharedString = string.Empty;
                    this.sharedStrings.TryGetValue(number, out sharedString);
                    
                    // Make column object
                    Column column = new Column(columnIndex);

                    // Add cell to row
                    Cell cell = new Cell(sharedString);
                    cell.Row = row;
                    cell.Column = column;
                    row.AddCell(cell);
                }

                // Add row to sheet
                sheet.AddRow(row);
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

        public IEnumerable<Sheet> Sheets
        {
            get
            {
                List<Sheet> sheetsToReturn = new List<Sheet>();

                IEnumerable<Sheet> visibleSheets = this.sheets.Where(s => s.Hidden == false);
                sheetsToReturn.AddRange(visibleSheets);

                if (this.options.IncludeHidden)
                {
                    IEnumerable<Sheet> hiddenSheets = this.sheets.Where(s => s.Hidden == true);
                    sheetsToReturn.AddRange(hiddenSheets);
                }

                return sheetsToReturn;    
            }
        }

        public WorkbookOptions Options
        {
            get { return this.options; }
            set { this.options = value; }
        }
    }
}
