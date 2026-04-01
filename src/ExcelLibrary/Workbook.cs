using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExcelLibrary;

public class Workbook
{
    private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string NS_OR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NS_ORW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    private readonly List<Sheet> sheets = new List<Sheet>();

    public string File { get; private set; }
    public Dictionary<int, string> SharedStrings { get; } = new Dictionary<int, string>();
    public Dictionary<int, NumberFormat> NumberFormats { get; } = new Dictionary<int, NumberFormat>();
    public WorkbookOptions Options { get; set; } = new WorkbookOptions();
    public int BaseYear { get; private set; } = 1900;

    public void Open(string file)
    {
        File = file;
        Open();
    }

    public void Open(string file, WorkbookOptions options)
    {
        File = file;
        Options = options;
        Open();
    }

    private void Open()
    {
        using (ZipArchive archive = ZipFile.OpenRead(File))
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

            // Read "xl/styles.xml" to get number formats
            ZipArchiveEntry stylesEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/styles.xml");
            LoadStyles(stylesEntry);

            // Optionally load all sheets
            if (Options.LoadSheets)
            {
                foreach (Sheet sheet in sheets)
                {
                    sheet.Open();
                }
            }
        }
    }

    private void LoadWorkbookXml(ZipArchiveEntry entry)
    {
        XDocument document = XDocument.Load(entry.Open());
        XElement root = document.Root;
        XNamespace ns = NS_MAIN;
        XNamespace r = NS_OR;

        XElement workbookPr = root.Element(ns + "workbookPr");
        XAttribute date1904 = workbookPr.Attribute("date1904");
        if (date1904 != null && date1904.Value == "1")
        {
            BaseYear = 1904;
        }

        foreach (XElement element in root.Element(ns + "sheets").Elements())
        {
            XAttribute id = element.Attribute(r + "id");
            XAttribute name = element.Attribute("name");
            XAttribute state = element.Attribute("state");

            bool hidden = false;
            if (state != null && state.Value == "hidden")
            {
                hidden = true;
            }

            Sheet sheet = new Sheet(name.Value, id.Value, hidden);
            sheet.Workbook = this;
            sheets.Add(sheet);
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
            {
                continue;
            }

            Sheet sheet = (from s in this.sheets
                           where s.Id == id.Value
                           select s).FirstOrDefault();

            // Get sheet file path
            Match match = Regex.Match(target.Value, @"worksheets/(.+)");
            string path = @"xl/worksheets/" + match.Groups[1].Value;
            sheet.Path = path.ToLower();
        }
    }

    private void LoadSharedStrings(ZipArchiveEntry entry)
    {
        // The xl/sharedStrings.xml file will be missing if there are no strings
        if (entry == null)
        {
            return;
        }

        XDocument document = XDocument.Load(entry.Open());
        XElement root = document.Root;
        XNamespace ns = NS_MAIN;
        int count = 0;

        foreach (XElement si in root.Elements(ns + "si"))
        {
            IEnumerable<XElement> ts = si.Descendants(ns + "t");
            string sum = string.Empty;
            foreach (XElement t in ts)
            {
                sum += t.Value;
            }

            SharedStrings.Add(count, sum);
            count++;
        }
    }

    private void LoadStyles(ZipArchiveEntry entry)
    {
        XDocument document = XDocument.Load(entry.Open());
        XNamespace ns = NS_MAIN;
        XElement cellXfs = document.Root.Element(ns + "cellXfs");
        int index = 0;

        foreach (XElement element in cellXfs.Elements())
        {
            XAttribute numFmtId = element.Attribute("numFmtId");
            if (numFmtId != null)
            {
                int numberFormatId = int.Parse(numFmtId.Value);
                switch (numberFormatId)
                {
                    case 0:
                        NumberFormats.Add(index, NumberFormat.General);
                        break;
                    case 2:
                        NumberFormats.Add(index, NumberFormat.Number);
                        break;
                    case 164:
                        NumberFormats.Add(index, NumberFormat.Currency);
                        break;
                    case 44:
                        NumberFormats.Add(index, NumberFormat.Accounting);
                        break;
                    case 14:
                        NumberFormats.Add(index, NumberFormat.Date);
                        break;
                    case 165:
                        NumberFormats.Add(index, NumberFormat.Time);
                        break;
                    case 49:
                        NumberFormats.Add(index, NumberFormat.Text);
                        break;
                    case 10:
                        NumberFormats.Add(index, NumberFormat.Percentage);
                        break;
                    case 13:
                        NumberFormats.Add(index, NumberFormat.Fraction);
                        break;
                    case 166:
                        NumberFormats.Add(index, NumberFormat.Custom);
                        break;
                    case 11:
                        NumberFormats.Add(index, NumberFormat.Scientific);
                        break;
                    default:
                        NumberFormats.Add(index, NumberFormat.Unsupported);
                        break;
                }
                index++;
            }
        }
    }

    public IEnumerable<Sheet> Sheets
    {
        get
        {
            if (Options.IncludeHidden)
            {
                return sheets.OrderBy(s => s.Id);
            }
            else
            {
                return sheets.Where(s => s.Hidden == false).OrderBy(s => s.Id);
            }
        }
    }

    public Sheet Sheet(string name)
    {
        return sheets.SingleOrDefault(s => s.Name == name);
    }
}
