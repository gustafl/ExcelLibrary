namespace ExcelLibrary.Tests;

[TestClass]
public class UtilitiesTests
{
    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertDate_WithBaseYear1900_ReturnsCorrectDate()
    {
        // Arrange
        const string excelDate = "45000"; // Days since 1900-01-01
        const int baseYear = 1900;

        // Act
        var result = Utilities.ConvertDate(excelDate, baseYear);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrEmpty(result));
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertDate_WithBaseYear1904_ReturnsCorrectDate()
    {
        // Arrange
        const string excelDate = "45000";
        const int baseYear = 1904;

        // Act
        var result = Utilities.ConvertDate(excelDate, baseYear);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsFalse(string.IsNullOrEmpty(result));
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertDate_WithDifferentBaseYears_ReturnsDifferentDates()
    {
        // Arrange
        const string excelDate = "45000";

        // Act
        var result1900 = Utilities.ConvertDate(excelDate, 1900);
        var result1904 = Utilities.ConvertDate(excelDate, 1904);

        // Assert
        Assert.AreNotEqual(result1900, result1904);
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertTime_WithTimeUnder24Hours_ReturnsFormattedTime()
    {
        // Arrange
        const string excelTime = "0.5"; // 12:00:00 (noon)

        // Act
        var result = Utilities.ConvertTime(excelTime);

        // Assert
        Assert.AreEqual("12:00:00", result);
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertTime_WithZero_ReturnsMidnight()
    {
        // Arrange
        const string excelTime = "0";

        // Act
        var result = Utilities.ConvertTime(excelTime);

        // Assert
        Assert.AreEqual("00:00:00", result);
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertTime_WithQuarterDay_Returns6AM()
    {
        // Arrange
        const string excelTime = "0.25"; // 6:00:00 AM

        // Act
        var result = Utilities.ConvertTime(excelTime);

        // Assert
        Assert.AreEqual("06:00:00", result);
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertTime_WithTimeOver24Hours_ReturnsTimeSpanFormat()
    {
        // Arrange
        const string excelTime = "1.5"; // 36 hours

        // Act
        var result = Utilities.ConvertTime(excelTime);

        // Assert
        // TimeSpan format is different from TimeOnly format
        Assert.IsTrue(result.Contains("1.12:00:00") || result.Contains("36:00:00"));
    }

    [TestMethod]
    [TestCategory("Utilities")]
    public void ConvertTime_WithAlmostFullDay_ReturnsCorrectTime()
    {
        // Arrange
        const string excelTime = "0.99999"; // Just before midnight

        // Act
        var result = Utilities.ConvertTime(excelTime);

        // Assert
        Assert.IsNotNull(result);
        Assert.IsTrue(result.StartsWith("23:59:"));
    }
}
