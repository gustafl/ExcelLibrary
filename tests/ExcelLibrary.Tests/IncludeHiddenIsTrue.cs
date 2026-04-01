namespace ExcelLibrary.Tests;

[TestClass]
public class IncludeHiddenIsTrue
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test1.xlsx");

    private const int ExpectedTotalSheetCount = 4;
    private const int ExpectedTotalRowCount = 5;
    private const int ExpectedTotalColumnCount = 4;
    private const int ExpectedTotalCellCount = 6;

    private Workbook workbook = null!;
    private WorkbookOptions options = null!;

    [TestInitialize]
    public void Initialize()
    {
        options = new WorkbookOptions { IncludeHidden = true, LoadSheets = true };
        workbook = new();
        workbook.Open(FILE, options);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Options_IncludeHidden_IsTrue()
    {
        // Assert
        Assert.IsTrue(options.IncludeHidden);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_WithIncludeHiddenTrue_ReturnsAllSheets()
    {
        // Act
        var sheets = workbook.Sheets;

        // Assert
        Assert.AreEqual(ExpectedTotalSheetCount, sheets.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Rows_WithIncludeHiddenTrue_ReturnsAllRows()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var rows = sheet.Rows;

        // Assert
        Assert.AreEqual(ExpectedTotalRowCount, rows.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Columns_WithIncludeHiddenTrue_ReturnsAllColumns()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var columns = sheet.Columns;

        // Assert
        Assert.AreEqual(ExpectedTotalColumnCount, columns.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cells_WithIncludeHiddenTrue_ReturnsAllCells()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cells = sheet.Cells;

        // Assert
        Assert.AreEqual(ExpectedTotalCellCount, cells.Count());
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_CellsWithIncludeHiddenTrue_ReturnsAllCells()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);
        Assert.IsNotNull(row);

        // Act
        var cellCount = row.Cells.Count();

        // Assert
        Assert.AreEqual(2, cellCount);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_WhenHidden_HiddenPropertyIsTrue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var row = sheet.Row(8);

        // Assert
        Assert.IsNotNull(row);
        Assert.IsTrue(row.Hidden);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_CellsWithIncludeHiddenTrue_ReturnsAllCells()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(3);

        // Act
        var cellCount = column.Cells.Count();

        // Assert
        Assert.IsNotNull(column);
        Assert.AreEqual(2, cellCount);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_WhenHidden_HiddenPropertyIsTrue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var column = sheet.Column(5);

        // Assert
        Assert.IsNotNull(column);
        Assert.IsTrue(column.Hidden);
    }
}
