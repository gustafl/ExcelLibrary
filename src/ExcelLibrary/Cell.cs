namespace ExcelLibrary;

public class Cell(string value)
{
    public required Row Row { get; set; }
    public required Column Column { get; set; }
    public string Value { get; set; } = value;
    public NumberFormat Format { get; set; }
}
