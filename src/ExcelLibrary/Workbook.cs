namespace ExcelLibrary;

/// <summary>
/// Represents an Excel workbook (.xlsx file) and provides access to its sheets, cells, and metadata.
/// </summary>
/// <example>
/// <code>
/// // Simple one-liner
/// using var workbook = Workbook.Open("data.xlsx");
/// var value = workbook.Sheet("Sheet1")?.Cell("A1")?.Value;
/// </code>
/// </example>
public class Workbook : IDisposable
{
    private const string NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string NS_OR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NS_ORW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    private static readonly FrozenDictionary<int, NumberFormat> NumberFormatMapping = new Dictionary<int, NumberFormat>
    {
        [0] = NumberFormat.General,
        [2] = NumberFormat.Number,
        [10] = NumberFormat.Percentage,
        [11] = NumberFormat.Scientific,
        [13] = NumberFormat.Fraction,
        [14] = NumberFormat.Date,
        [44] = NumberFormat.Accounting,
        [49] = NumberFormat.Text,
        [164] = NumberFormat.Currency,
        [165] = NumberFormat.Time,
        [166] = NumberFormat.Custom
    }.ToFrozenDictionary();

    private readonly List<Sheet> sheets = [];
    private readonly Dictionary<int, string> sharedStrings = [];
    private readonly Dictionary<int, NumberFormat> numberFormats = [];

    /// <summary>
    /// Gets the file path of the currently opened workbook.
    /// </summary>
    public string? File { get; private set; }

    /// <summary>
    /// Gets the shared strings table used by the workbook for string cell values.
    /// </summary>
    public IReadOnlyDictionary<int, string> SharedStrings => sharedStrings;

    /// <summary>
    /// Gets the number formats defined in the workbook, mapped by style index.
    /// </summary>
    public IReadOnlyDictionary<int, NumberFormat> NumberFormats => numberFormats;

    /// <summary>
    /// Gets the options controlling how the workbook is loaded and accessed.
    /// </summary>
    public WorkbookOptions Options { get; private set; } = new();

    /// <summary>
    /// Gets the base year for date calculations. Returns 1904 for Mac-created files, otherwise 1900.
    /// </summary>
    public int BaseYear { get; private set; } = 1900;

    /// <summary>
    /// Opens a workbook from the specified file path.
    /// </summary>
    /// <param name="file">The path to the .xlsx file.</param>
    /// <returns>A new <see cref="Workbook"/> instance with data loaded.</returns>
    /// <example>
    /// <code>
    /// using var workbook = Workbook.Open("data.xlsx");
    /// var value = workbook.Sheet("Sheet1")?.Cell("A1")?.Value;
    /// </code>
    /// </example>
    public static Workbook Open(string file) => Open(file, new WorkbookOptions());

    /// <summary>
    /// Opens a workbook from the specified file path with custom options.
    /// </summary>
    /// <param name="file">The path to the .xlsx file.</param>
    /// <param name="options">The options controlling how the workbook is loaded.</param>
    /// <returns>A new <see cref="Workbook"/> instance with data loaded.</returns>
    public static Workbook Open(string file, WorkbookOptions options)
    {
        var workbook = new Workbook { File = file, Options = options };
        workbook.Load();
        return workbook;
    }

    private void Load()
    {
        if (File is null) return;
        using var archive = ZipFile.OpenRead(File);

        // Collect sheet metadata from both XML files before creating Sheet objects
        var sheetMetadata = new Dictionary<string, SheetMetadata>();

        // Read "xl/workbook.xml" to get sheet names, ids, and hidden state
        var sheetsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/workbook.xml");
        if (sheetsEntry is not null)
            LoadWorkbookXml(sheetsEntry, sheetMetadata);

        // Read "xl/_rels/workbook.xml.rels" to get sheet paths
        var sheetPathsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/_rels/workbook.xml.rels");
        if (sheetPathsEntry is not null)
            LoadWorkbookXmlRels(sheetPathsEntry, sheetMetadata);

        // Now create Sheet objects with complete metadata
        foreach (var metadata in sheetMetadata.Values)
        {
            var sheet = new Sheet(metadata.Name, metadata.Id, metadata.Hidden)
            {
                Workbook = this,
                Path = metadata.Path ?? string.Empty
            };
            sheets.Add(sheet);
        }

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
            if (Options.ParallelLoadSheets)
                Parallel.ForEach(sheets, sheet => sheet.Open());
            else
                foreach (var sheet in sheets)
                    sheet.Open();
        }
    }

    /// <summary>
    /// Temporary container for sheet metadata during workbook loading.
    /// </summary>
    private sealed record SheetMetadata(string Name, string? Id, bool Hidden)
    {
        public string? Path { get; set; }
    }

    private void LoadWorkbookXml(ZipArchiveEntry entry, Dictionary<string, SheetMetadata> sheetMetadata)
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
            if (name is null || id is null) continue;

            bool hidden = element.Attribute("state") is { Value: "hidden" };
            sheetMetadata[id] = new SheetMetadata(name, id, hidden);
        }
    }

    private void LoadWorkbookXmlRels(ZipArchiveEntry entry, Dictionary<string, SheetMetadata> sheetMetadata)
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

            if (!sheetMetadata.TryGetValue(id, out var metadata)) continue;

            // Get sheet file path
            var match = Regex.Match(target, @"worksheets/(.+)");
            metadata.Path = $"xl/worksheets/{match.Groups[1].Value}".ToLower();
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
            sharedStrings.Add(count++, sum);
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
                int formatId = int.Parse(numFmtId.Value);
                var format = NumberFormatMapping.GetValueOrDefault(formatId, NumberFormat.Unsupported);
                numberFormats.Add(index++, format);
            }
        }
    }

    /// <summary>
    /// Gets the sheets in the workbook. Hidden sheets are excluded unless <see cref="WorkbookOptions.IncludeHidden"/> is <c>true</c>.
    /// </summary>
    public IEnumerable<Sheet> Sheets =>
        Options.IncludeHidden
            ? sheets.OrderBy(s => s.Id)
            : sheets.Where(s => !s.Hidden).OrderBy(s => s.Id);

    /// <summary>
    /// Gets a sheet by its name.
    /// </summary>
    /// <param name="name">The name of the sheet.</param>
    /// <returns>The sheet with the specified name, or <c>null</c> if not found.</returns>
    public Sheet? Sheet(string name) => sheets.Find(s => string.Equals(s.Name, name, StringComparison.Ordinal));

    /// <summary>
    /// Releases all resources used by this workbook.
    /// </summary>
    public void Dispose()
    {
        // Currently no unmanaged resources to dispose.
        // This method exists to support the IDisposable pattern for future extensibility.
        GC.SuppressFinalize(this);
    }
}