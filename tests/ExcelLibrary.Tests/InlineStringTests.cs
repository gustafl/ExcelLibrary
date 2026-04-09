namespace ExcelLibrary.Tests;

using System.IO;
using System.IO.Compression;
using System.Text;

[TestClass]
public class InlineStringTests
{
    private static readonly string TestFile = Path.Combine(AppContext.BaseDirectory, "Input", "inline_strings.xlsx");

    [ClassInitialize]
    public static void CreateTestFile(TestContext context)
    {
        // Create a minimal .xlsx file with inline strings
        // .xlsx files are ZIP archives containing XML files
        
        var dir = Path.GetDirectoryName(TestFile)!;
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        if (File.Exists(TestFile))
            File.Delete(TestFile);

        using var zip = ZipFile.Open(TestFile, ZipArchiveMode.Create);

        // [Content_Types].xml
        AddEntry(zip, "[Content_Types].xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            </Types>
            """);

        // _rels/.rels
        AddEntry(zip, "_rels/.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);

        // xl/workbook.xml
        AddEntry(zip, "xl/workbook.xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="InlineSheet" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """);

        // xl/_rels/workbook.xml.rels
        AddEntry(zip, "xl/_rels/workbook.xml.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """);

        // xl/worksheets/sheet1.xml - Contains inline strings (t="inlineStr" with <is><t>)
        AddEntry(zip, "xl/worksheets/sheet1.xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
                        <c r="B1" t="inlineStr"><is><t>World</t></is></c>
                    </row>
                    <row r="2">
                        <c r="A2" t="inlineStr"><is><t>Inline</t></is></c>
                        <c r="B2" t="inlineStr"><is><t>String</t></is></c>
                    </row>
                    <row r="3">
                        <c r="A3"><v>42</v></c>
                        <c r="B3" t="inlineStr"><is><t>Mixed</t></is></c>
                    </row>
                </sheetData>
            </worksheet>
            """);
    }

    private static void AddEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name);
        using var stream = entry.Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        if (File.Exists(TestFile))
            File.Delete(TestFile);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Open_WithInlineStrings_ReturnsWorkbookInstance()
    {
        // Act
        using var workbook = Workbook.Open(TestFile);

        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Sheet_WithInlineStrings_ReturnsCorrectSheet()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);

        // Act
        var sheet = workbook.Sheet("InlineSheet");

        // Assert
        Assert.IsNotNull(sheet);
        Assert.AreEqual("InlineSheet", sheet.Name);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Cell_WithInlineString_ReturnsCorrectValue()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");

        // Act
        var cell = sheet.Cell("A1");

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Hello", cell.Value);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Cell_WithInlineString_ByRowAndColumn_ReturnsCorrectValue()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");

        // Act
        var cell = sheet.Cell(1, 2);

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("World", cell.Value);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Cells_WithInlineStrings_ReturnsAllCells()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");

        // Act
        var cells = sheet.Cells.ToList();

        // Assert
        Assert.AreEqual(6, cells.Count);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Row_WithInlineStrings_ReturnsCorrectCells()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");
        var row = sheet.Row(2);

        // Act
        var cells = row.Cells.ToList();

        // Assert
        Assert.AreEqual(2, cells.Count);
        Assert.AreEqual("Inline", cells[0].Value);
        Assert.AreEqual("String", cells[1].Value);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Sheet_WithMixedContent_ReturnsAllValues()
    {
        // Arrange - Row 3 has a numeric value and an inline string
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");

        // Act
        var numericCell = sheet.Cell("A3");
        var inlineCell = sheet.Cell("B3");

        // Assert
        Assert.IsNotNull(numericCell);
        Assert.AreEqual("42", numericCell.Value);
        Assert.IsNotNull(inlineCell);
        Assert.AreEqual("Mixed", inlineCell.Value);
    }

    [TestMethod]
    [TestCategory("InlineString")]
    public void Rows_WithInlineStrings_ReturnsCorrectCount()
    {
        // Arrange
        using var workbook = Workbook.Open(TestFile);
        var sheet = workbook.Sheet("InlineSheet");

        // Act
        var rows = sheet.Rows.ToList();

        // Assert
        Assert.AreEqual(3, rows.Count);
    }
}
