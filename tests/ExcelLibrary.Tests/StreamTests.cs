namespace ExcelLibrary.Tests;

using System.IO;

[TestClass]
public class StreamTests
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test1.xlsx");

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithStream_ReturnsWorkbookInstance()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);

        // Act
        using var workbook = Workbook.Open(stream);

        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithStream_FilePropertyIsNull()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);

        // Act
        using var workbook = Workbook.Open(stream);

        // Assert
        Assert.IsNull(workbook.File);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithStreamAndOptions_ReturnsWorkbookInstance()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);
        var options = new WorkbookOptions { IncludeHidden = true };

        // Act
        using var workbook = Workbook.Open(stream, options);

        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
        Assert.IsTrue(workbook.Options.IncludeHidden);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_WithStream_ReturnsAllSheets()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);

        // Act
        using var workbook = Workbook.Open(stream);
        var sheets = workbook.Sheets.ToList();

        // Assert
        Assert.AreEqual(3, sheets.Count); // 3 visible sheets
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_WithStream_ReturnsCorrectValue()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);

        // Act
        using var workbook = Workbook.Open(stream);
        var cell = workbook.Sheet("Sheet1")?.Cell("B2");

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithMemoryStream_ReturnsWorkbookInstance()
    {
        // Arrange - Load file into memory stream
        var bytes = File.ReadAllBytes(FILE);
        using var memoryStream = new MemoryStream(bytes);

        // Act
        using var workbook = Workbook.Open(memoryStream);

        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_WithMemoryStream_ReturnsCorrectValue()
    {
        // Arrange
        var bytes = File.ReadAllBytes(FILE);
        using var memoryStream = new MemoryStream(bytes);

        // Act
        using var workbook = Workbook.Open(memoryStream);
        var cell = workbook.Sheet("Sheet1")?.Cell("B2");

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithStreamAndLazyLoading_SheetsCanBeLoadedLater()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);
        var options = new WorkbookOptions { LoadSheets = false };

        // Act
        using var workbook = Workbook.Open(stream, options);
        var sheet = workbook.Sheet("Sheet1");

        // Assert - Sheet exists but has no data yet
        Assert.IsNotNull(sheet);
        Assert.AreEqual(0, sheet.Rows.Count());

        // Load the sheet
        sheet.Open();

        // Now it has data
        Assert.IsTrue(sheet.Rows.Any());
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithStreamAndParallelLoading_LoadsAllSheets()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);
        var options = new WorkbookOptions { ParallelLoadSheets = true };

        // Act
        using var workbook = Workbook.Open(stream, options);
        var sheetsWithData = workbook.Sheets.Where(s => s.Rows.Any()).ToList();

        // Assert
        Assert.AreEqual(2, sheetsWithData.Count); // Sheet1 and Sheet2 have data
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void StreamAndFile_ReturnSameData()
    {
        // Arrange
        using var stream = File.OpenRead(FILE);

        // Act
        using var streamWorkbook = Workbook.Open(stream);
        using var fileWorkbook = Workbook.Open(FILE);

        var streamCells = streamWorkbook.Sheets
            .SelectMany(s => s.Cells)
            .Select(c => c.Value)
            .OrderBy(v => v)
            .ToList();

        var fileCells = fileWorkbook.Sheets
            .SelectMany(s => s.Cells)
            .Select(c => c.Value)
            .OrderBy(v => v)
            .ToList();

        // Assert
        CollectionAssert.AreEqual(fileCells, streamCells);
    }
}
