namespace ExcelLibrary;

/// <summary>
/// Configuration options for opening and accessing an Excel workbook.
/// </summary>
public record WorkbookOptions
{
    /// <summary>
    /// Gets or sets whether hidden sheets, rows, and columns should be included when accessing workbook data.
    /// Default is <c>false</c>.
    /// </summary>
    public bool IncludeHidden { get; init; }

    /// <summary>
    /// Gets or sets whether all sheets should be loaded immediately when opening the workbook.
    /// When <c>false</c>, sheets are loaded on-demand via <see cref="Sheet.Open"/>.
    /// Default is <c>true</c>.
    /// </summary>
    public bool LoadSheets { get; init; } = true;
}
