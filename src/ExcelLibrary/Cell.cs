namespace ExcelLibrary;

public class Cell(string value)
{
    public Row Row { get; set; }
    public Column Column { get; set; }
    public string Value { get; set; } = value;
    public NumberFormat Format { get; set; }
}
