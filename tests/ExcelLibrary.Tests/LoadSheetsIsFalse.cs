namespace ExcelLibrary.Tests;

[TestClass]
public class LoadSheetsIsFalse
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test1.xlsx");

    private Workbook workbook = null!;
    private WorkbookOptions options = null!;

    [TestInitialize]
    public void Initialize()
    {
        options = new WorkbookOptions { IncludeHidden = true, LoadSheets = false };
        workbook = Workbook.Open(FILE, options);
    }

    [TestMethod]
    [TestCategory("Workbook")]
    public void Options_LoadSheets_IsFalse()
    {
        // Assert
        Assert.IsFalse(options.LoadSheets);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Sheet_Open_ReturnsSheetInstance()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        sheet.Open();

        // Assert
        Assert.IsInstanceOfType<Sheet>(sheet);
    }

    [TestMethod]
    [TestCategory("Sheet")]
    public void Rows_BeforeSheetOpen_ReturnsEmptyCollection()
    {
        // Arrange
        var sheet = workbook.Sheet("Sheet1");

        // Act
        var rows = sheet.Rows;

        // Assert
        Assert.AreEqual(0, rows.Count());
    }
}
