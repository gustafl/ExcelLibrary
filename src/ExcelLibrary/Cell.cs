namespace ExcelLibrary;

/// <summary>
/// Represents a cell in an Excel worksheet.
/// </summary>
/// <param name="value">The string value of the cell.</param>
public class Cell(string value)
{
    /// <summary>
    /// Gets the row containing this cell.
    /// </summary>
    public required Row Row { get; init; }

    /// <summary>
    /// Gets the column containing this cell.
    /// </summary>
    public required Column Column { get; init; }

    /// <summary>
    /// Gets the cell's value as a string.
    /// </summary>
    public string Value { get; init; } = value;

    /// <summary>
    /// Gets the number format applied to this cell.
    /// </summary>
    public NumberFormat Format { get; init; } = NumberFormat.General;
}
