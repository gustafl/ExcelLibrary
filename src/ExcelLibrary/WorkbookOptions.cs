namespace ExcelLibrary;

public record WorkbookOptions
{
    public bool IncludeHidden { get; init; }
    public bool LoadSheets { get; init; } = true;
}
