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

    private readonly List<Sheet> sheets = [];

    public string? File { get; private set; }
    public Dictionary<int, string> SharedStrings { get; } = [];
    public Dictionary<int, NumberFormat> NumberFormats { get; } = [];
    public WorkbookOptions Options { get; set; } = new();
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
        if (File is null) return;
        using var archive = ZipFile.OpenRead(File);

        // Read "xl/workbook.xml" to get sheet names and ids
        var sheetsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/workbook.xml");
        if (sheetsEntry is not null)
            LoadWorkbookXml(sheetsEntry);

        // Read "xl/_rels/workbook.xml.rels" to get sheet paths
        var sheetPathsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/_rels/workbook.xml.rels");
        if (sheetPathsEntry is not null)
            LoadWorkbookXmlRels(sheetPathsEntry);

        // Read "xl/sharedStrings.xml" to get shared strings
        var sharedStringsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/sharedStrings.xml");
        if (sharedStringsEntry is not null)
            LoadSharedStrings(sharedStringsEntry);

        // Read "xl/styles.xml" to get number formats
        var stylesEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/styles.xml");
        if (stylesEntry is not null)
            LoadStyles(stylesEntry);

        // Optionally load all sheets
        if (Options.LoadSheets)
        {
            foreach (var sheet in sheets)
                sheet.Open();
        }
    }

    private void LoadWorkbookXml(ZipArchiveEntry entry)
    {
        var document = XDocument.Load(entry.Open());
        var root = document.Root;
        if (root is null) return;

        XNamespace ns = NS_MAIN;
        XNamespace r = NS_OR;

        var workbookPr = root.Element(ns + "workbookPr");
        if (workbookPr?.Attribute("date1904") is { Value: "1" })
            BaseYear = 1904;

        var sheetsElement = root.Element(ns + "sheets");
        if (sheetsElement is null) return;

        foreach (var element in sheetsElement.Elements())
        {
            var id = element.Attribute(r + "id")?.Value;
            var name = element.Attribute("name")?.Value;
            if (name is null) continue;

            bool hidden = element.Attribute("state") is { Value: "hidden" };
            var sheet = new Sheet(name, id, hidden) { Workbook = this };
            sheets.Add(sheet);
        }
    }

    private void LoadWorkbookXmlRels(ZipArchiveEntry entry)
    {
        var document = XDocument.Load(entry.Open());
        var root = document.Root;
        if (root is null) return;

        XNamespace ns = NS_PR;

        foreach (var element in root.Elements(ns + "Relationship"))
        {
            var type = element.Attribute("Type");
            if (type?.Value != NS_ORW)
                continue;

            var id = element.Attribute("Id")?.Value;
            var target = element.Attribute("Target")?.Value;
            if (id is null || target is null) continue;

            var sheet = sheets.FirstOrDefault(s => s.Id == id);
            if (sheet is null) continue;

            // Get sheet file path
            var match = Regex.Match(target, @"worksheets/(.+)");
            sheet.Path = $"xl/worksheets/{match.Groups[1].Value}".ToLower();
        }
    }

    private void LoadSharedStrings(ZipArchiveEntry entry)
    {
        var document = XDocument.Load(entry.Open());
        var root = document.Root;
        if (root is null) return;

        XNamespace ns = NS_MAIN;
        int count = 0;

        foreach (var si in root.Elements(ns + "si"))
        {
            string sum = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
            SharedStrings.Add(count++, sum);
        }
    }

    private void LoadStyles(ZipArchiveEntry entry)
    {
        var document = XDocument.Load(entry.Open());
        XNamespace ns = NS_MAIN;
        var cellXfs = document.Root?.Element(ns + "cellXfs");
        if (cellXfs is null) return;

        int index = 0;

        foreach (var element in cellXfs.Elements())
        {
            if (element.Attribute("numFmtId") is { } numFmtId)
            {
                int numberFormatId = int.Parse(numFmtId.Value);
                NumberFormats.Add(index++, numberFormatId switch
                {
                    0 => NumberFormat.General,
                    2 => NumberFormat.Number,
                    164 => NumberFormat.Currency,
                    44 => NumberFormat.Accounting,
                    14 => NumberFormat.Date,
                    165 => NumberFormat.Time,
                    49 => NumberFormat.Text,
                    10 => NumberFormat.Percentage,
                    13 => NumberFormat.Fraction,
                    166 => NumberFormat.Custom,
                    11 => NumberFormat.Scientific,
                    _ => NumberFormat.Unsupported
                });
            }
        }
    }

    public IEnumerable<Sheet> Sheets =>
        Options.IncludeHidden
            ? sheets.OrderBy(s => s.Id)
            : sheets.Where(s => !s.Hidden).OrderBy(s => s.Id);

    public Sheet? Sheet(string name) => sheets.SingleOrDefault(s => s.Name == name);
}
