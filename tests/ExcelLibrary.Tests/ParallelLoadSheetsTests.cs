namespace ExcelLibrary.Tests;

[TestClass]
public class ParallelLoadSheetsTests
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test1.xlsx");

    [TestMethod]
    [TestCategory("Workbook")]
    public void Options_ParallelLoadSheets_DefaultIsFalse()
    {
        // Arrange
        var options = new WorkbookOptions();

        // Assert
        Assert.IsFalse(options.ParallelLoadSheets);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Options_ParallelLoadSheets_CanBeSetToTrue()
    {
        // Arrange
        var options = new WorkbookOptions { ParallelLoadSheets = true };

        // Assert
        Assert.IsTrue(options.ParallelLoadSheets);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithParallelLoadSheets_ReturnsWorkbookInstance()
    {
        // Arrange
        var options = new WorkbookOptions { ParallelLoadSheets = true };

        // Act
        using var workbook = Workbook.Open(FILE, options);

        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_WithParallelLoadSheets_ReturnsAllSheets()
    {
        // Arrange
        var options = new WorkbookOptions { ParallelLoadSheets = true, IncludeHidden = true };

        // Act
        using var workbook = Workbook.Open(FILE, options);
        var sheets = workbook.Sheets.ToList();

        // Assert
        Assert.AreEqual(4, sheets.Count);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_WithParallelLoadSheets_AllSheetsHaveData()
    {
        // Arrange
        var options = new WorkbookOptions { ParallelLoadSheets = true };

        // Act
        using var workbook = Workbook.Open(FILE, options);
        var sheetsWithRows = workbook.Sheets.Where(s => s.Rows.Any()).ToList();

        // Assert - Sheet1 and Sheet2 have data, Sheet3 is empty
        Assert.AreEqual(2, sheetsWithRows.Count);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_WithParallelLoadSheets_ReturnsCorrectValue()
    {
        // Arrange
        var options = new WorkbookOptions { ParallelLoadSheets = true };

        // Act
        using var workbook = Workbook.Open(FILE, options);
        var sheet = workbook.Sheet("Sheet1");
        var cell = sheet?.Cell("B2");

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithParallelLoadSheetsAndLoadSheetsFalse_DoesNotLoadSheets()
    {
        // Arrange - ParallelLoadSheets should have no effect when LoadSheets is false
        var options = new WorkbookOptions { ParallelLoadSheets = true, LoadSheets = false };

        // Act
        using var workbook = Workbook.Open(FILE, options);
        var sheet = workbook.Sheet("Sheet1");

        // Assert - Sheet exists but has no data (not loaded)
        Assert.IsNotNull(sheet);
        Assert.AreEqual(0, sheet.Rows.Count());
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_ParallelAndSequential_ReturnSameData()
    {
        // Arrange
        var parallelOptions = new WorkbookOptions { ParallelLoadSheets = true };
        var sequentialOptions = new WorkbookOptions { ParallelLoadSheets = false };

        // Act
        using var parallelWorkbook = Workbook.Open(FILE, parallelOptions);
        using var sequentialWorkbook = Workbook.Open(FILE, sequentialOptions);

        var parallelCells = parallelWorkbook.Sheets
            .SelectMany(s => s.Cells)
            .Select(c => c.Value)
            .OrderBy(v => v)
            .ToList();

        var sequentialCells = sequentialWorkbook.Sheets
            .SelectMany(s => s.Cells)
            .Select(c => c.Value)
            .OrderBy(v => v)
            .ToList();

        // Assert - Both loading methods should return identical data
        CollectionAssert.AreEqual(sequentialCells, parallelCells);
    }
}
