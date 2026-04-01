namespace ExcelLibrary;

/// <summary>
/// Represents a cell in an Excel worksheet.
/// </summary>
/// <param name="value">The string value of the cell.</param>
public class Cell(string value)
{
    /// <summary>
    /// Gets or sets the row containing this cell.
    /// </summary>
    public required Row Row { get; set; }

    /// <summary>
    /// Gets or sets the column containing this cell.
    /// </summary>
    public required Column Column { get; set; }

    /// <summary>
    /// Gets or sets the cell's value as a string.
    /// </summary>
    public string Value { get; set; } = value;

    /// <summary>
    /// Gets or sets the number format applied to this cell.
    /// </summary>
    public NumberFormat Format { get; set; }
}
