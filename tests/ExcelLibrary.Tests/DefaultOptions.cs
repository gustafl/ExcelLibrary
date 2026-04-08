namespace ExcelLibrary.Tests;

[TestClass]
public class DefaultOptions
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test1.xlsx");

    private const int ExpectedVisibleSheetCount = 3;
    private const int ExpectedSharedStringCount = 35;
    private const int ExpectedVisibleRowCount = 4;
    private const int ExpectedVisibleColumnCount = 3;
    private const int ExpectedVisibleCellCount = 4;

    private Workbook workbook = null!;

    [TestInitialize]
    public void Initialize()
    {
        workbook = Workbook.Open(FILE);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Open_WithDefaultOptions_ReturnsWorkbookInstance()
    {
        // Assert
        Assert.IsInstanceOfType<Workbook>(workbook);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void File_AfterOpen_ReturnsFilePath()
    {
        // Assert
        Assert.AreEqual(FILE, workbook.File);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Sheets_WithDefaultOptions_ReturnsOnlyVisibleSheets()
    {
        // Act
        var sheets = workbook.Sheets;

        // Assert
        Assert.AreEqual(ExpectedVisibleSheetCount, sheets.Count());
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void SharedStrings_AfterOpen_ReturnsExpectedCount()
    {
        // Assert
        Assert.AreEqual(ExpectedSharedStringCount, workbook.SharedStrings.Count);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_ByName_ReturnsCorrectName()
    {
        // Act
        var sheet = workbook.Sheet("Sheet1");

        // Assert
        Assert.AreEqual("Sheet1", sheet.Name);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_ByName_ReturnsCorrectId()
    {
        // Act
        var sheet = workbook.Sheet("Sheet1");

        // Assert
        Assert.AreEqual("rId1", sheet.Id);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_ByName_ReturnsCorrectPath()
    {
        // Act
        var sheet = workbook.Sheet("Sheet1");

        // Assert
        Assert.AreEqual("xl/worksheets/sheet1.xml", sheet.Path.ToLower());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_WhenHidden_HiddenPropertyIsTrue()
    {
        // Act
        var sheet = workbook.Sheet("Sheet4");

        // Assert
        Assert.IsTrue(sheet.Hidden);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_Workbook_ReturnsSameInstance()
    {
        // Act
        var sheet = workbook.Sheet("Sheet1");

        // Assert
        Assert.AreSame(workbook, sheet.Workbook);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Rows_WithDefaultOptions_ReturnsOnlyVisibleRows()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var rows = sheet.Rows;

        // Assert
        Assert.AreEqual(ExpectedVisibleRowCount, rows.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Columns_WithDefaultOptions_ReturnsOnlyVisibleColumns()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var columns = sheet.Columns;

        // Assert
        Assert.AreEqual(ExpectedVisibleColumnCount, columns.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Row_ByIndex_ReturnsCorrectCellValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var text = row.Cell(2)?.Value;

        // Assert
        Assert.IsNotNull(row);
        Assert.AreEqual("Banana", text);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Column_ByIndex_ReturnsCorrectCellValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(2);

        // Act
        var text = column.Cell(2)?.Value;

        // Assert
        Assert.IsNotNull(column);
        Assert.AreEqual("Banana", text);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cells_WithDefaultOptions_ReturnsOnlyVisibleCells()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cells = sheet.Cells;

        // Assert
        Assert.AreEqual(ExpectedVisibleCellCount, cells.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_ByRowAndColumnIndex_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell(2, 2);

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_ByName_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell("B2");

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cells_InRowsAndColumns_HaveSameCount()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var rowCellCount = sheet.Rows.SelectMany(r => r.Cells).Count();
        var columnCellCount = sheet.Columns.SelectMany(c => c.Cells).Count();

        // Assert
        Assert.AreEqual(rowCellCount, columnCellCount);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Row_WhenNotExists_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var row = sheet.Row(100);

        // Assert
        Assert.IsNull(row);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Column_WhenNotExists_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var column = sheet.Column(100);

        // Assert
        Assert.IsNull(column);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Row_WhenHiddenAndIncludeHiddenFalse_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var row = sheet.Row(8);

        // Assert
        Assert.IsNull(row);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Column_WhenHiddenAndIncludeHiddenFalse_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var column = sheet.Column(5);

        // Assert
        Assert.IsNull(column);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cells_OnEmptySheet_ReturnsEmptyCollection()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet3");

        // Act
        var cells = sheet.Cells;

        // Assert
        Assert.AreEqual(0, cells.Count());
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_WhenNotExistsByIndex_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell(100, 100);

        // Assert
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_WhenNotExistsByName_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell("A100");

        // Assert
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_InHiddenRow_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell(8, 3);

        // Assert
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_InHiddenColumn_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var cell = sheet.Cell(2, 5);

        // Assert
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_Index_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var index = row.Index;

        // Assert
        Assert.IsNotNull(row);
        Assert.AreEqual(2, index);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_Sheet_ReturnsSameInstance()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var rowSheet = row.Sheet;

        // Assert
        Assert.IsNotNull(row);
        Assert.AreSame(sheet, rowSheet);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_Cells_ReturnsCorrectCount()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var cellCount = row.Cells.Count();

        // Assert
        Assert.IsNotNull(row);
        Assert.AreEqual(1, cellCount);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_CellByColumnIndex_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var cell = row.Cell(2);

        // Assert
        Assert.IsNotNull(row);
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_CellInHiddenColumn_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act
        var cell = row.Cell(5);

        // Assert
        Assert.IsNotNull(row);
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_Index_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(2);

        // Act
        var index = column.Index;

        // Assert
        Assert.IsNotNull(column);
        Assert.AreEqual(2, index);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_Sheet_ReturnsSameInstance()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(2);

        // Act
        var columnSheet = column.Sheet;

        // Assert
        Assert.IsNotNull(column);
        Assert.AreSame(sheet, columnSheet);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_Cells_ReturnsCorrectCount()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(3);

        // Act
        var cellCount = column.Cells.Count();

        // Assert
        Assert.IsNotNull(column);
        Assert.AreEqual(1, cellCount);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_CellByRowIndex_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(2);

        // Act
        var cell = column.Cell(2);

        // Assert
        Assert.IsNotNull(column);
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", cell.Value);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_CellInHiddenRow_ReturnsNull()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(3);

        // Act
        var cell = column.Cell(8);

        // Assert
        Assert.IsNotNull(column);
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Cell")]
    public void Cell_Value_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);
        var cell = row.Cell(2);

        // Act
        var value = cell.Value;

        // Assert
        Assert.IsNotNull(row);
        Assert.IsNotNull(cell);
        Assert.AreEqual("Banana", value);
    }

    [TestMethod]
    [TestCategory("Cell")]
    public void Cell_Row_ReturnsSameInstance()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var row = sheet.Row(2);
        var cell = row.Cell(2);

        // Act
        var cellRow = cell.Row;

        // Assert
        Assert.IsNotNull(row);
        Assert.IsNotNull(cell);
        Assert.AreSame(row, cellRow);
    }

    [TestMethod]
    [TestCategory("Cell")]
    public void Cell_Column_ReturnsSameInstance()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");
        var column = sheet.Column(2);
        var cell = column.Cell(2);

        // Act
        var cellColumn = cell.Column;

        // Assert
        Assert.IsNotNull(cell);
        Assert.IsNotNull(column);
        Assert.AreSame(column, cellColumn);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void BaseYear_WithDefaultOptions_Returns1900()
    {
        // Act
        var baseYear = workbook.BaseYear;

        // Assert
        Assert.AreEqual(1900, baseYear);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Options_AfterOpen_ReturnsDefaultOptions()
    {
        // Act
        var options = workbook.Options;

        // Assert
        Assert.IsNotNull(options);
        Assert.IsFalse(options.IncludeHidden);
        Assert.IsTrue(options.LoadSheets);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void NumberFormats_AfterOpen_ReturnsNonEmptyDictionary()
    {
        // Act
        var numberFormats = workbook.NumberFormats;

        // Assert
        Assert.IsNotNull(numberFormats);
        Assert.IsTrue(numberFormats.Count > 0);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_WhenNameNotExists_ReturnsNull()
    {
        // Act
        var sheet = workbook.Sheet("NonExistentSheet");

        // Assert
        Assert.IsNull(sheet);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_Hidden_WhenVisible_ReturnsFalse()
    {
        // Arrange
        var options = new WorkbookOptions { IncludeHidden = true };
        var wb = Workbook.Open(FILE, options);
        var sheet = wb.Sheet("Sheet1");
        var row = sheet.Row(2);

        // Act & Assert
        Assert.IsNotNull(row);
        Assert.IsFalse(row.Hidden);
    }

    [TestMethod]
    [TestCategory("Row")]
    public void Row_Hidden_WhenHidden_ReturnsTrue()
    {
        // Arrange
        var options = new WorkbookOptions { IncludeHidden = true };
        var wb = Workbook.Open(FILE, options);
        var sheet = wb.Sheet("Sheet1");
        var row = sheet.Row(8);

        // Act & Assert
        Assert.IsNotNull(row);
        Assert.IsTrue(row.Hidden);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_Hidden_WhenVisible_ReturnsFalse()
    {
        // Arrange
        var options = new WorkbookOptions { IncludeHidden = true };
        var wb = Workbook.Open(FILE, options);
        var sheet = wb.Sheet("Sheet1");
        var column = sheet.Column(2);

        // Act & Assert
        Assert.IsNotNull(column);
        Assert.IsFalse(column.Hidden);
    }

    [TestMethod]
    [TestCategory("Column")]
    public void Column_Hidden_WhenHidden_ReturnsTrue()
    {
        // Arrange
        var options = new WorkbookOptions { IncludeHidden = true };
        var wb = Workbook.Open(FILE, options);
        var sheet = wb.Sheet("Sheet1");
        var column = sheet.Column(5);

        // Act & Assert
        Assert.IsNotNull(column);
        Assert.IsTrue(column.Hidden);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Cell_ByMultiLetterColumnName_ReturnsCorrectValue()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet2");

        // Act
        var cell = sheet.Cell("AA1");

        // Assert - AA1 may or may not exist, but should not throw
        // If no cell exists at AA1, it should return null
        Assert.IsNull(cell);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_WhenVisible_HiddenPropertyIsFalse()
    {
        // Act
        var sheet = workbook.Sheet("Sheet1");

        // Assert
        Assert.IsFalse(sheet.Hidden);
    }
}
