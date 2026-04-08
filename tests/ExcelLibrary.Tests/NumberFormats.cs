namespace ExcelLibrary.Tests;

[TestClass]
public class NumberFormats
{
    private static readonly string FILE = Path.Combine(AppContext.BaseDirectory, "Input", "test2.xlsx");

    private const decimal ExpectedValue = 123.45m;

    private Workbook workbook = null!;
    private Sheet sheet = null!;
    private Column column = null!;

    [TestInitialize]
    public void Initialize()
    {
        workbook = Workbook.Open(FILE);
        sheet = workbook.Sheet("Sheet1");
        column = sheet.Column(2);
    }

    [TestMethod]
    [TestCategory("NumberFormats")]
    [DataRow(1, DisplayName = "General format")]
    [DataRow(2, DisplayName = "Number format")]
    [DataRow(3, DisplayName = "Currency format")]
    [DataRow(4, DisplayName = "Accounting format")]
    public void Cell_WithNumberFormat_ReturnsCorrectValue(int rowIndex)
    {
        // Arrange
        var cell = column.Cell(rowIndex);

        // Act
        var val = cell.Value.Replace(".", ",");
        var number = decimal.Parse(val);

        // Assert
        Assert.IsNotNull(cell);
        Assert.AreEqual(ExpectedValue, number);
    }
}
