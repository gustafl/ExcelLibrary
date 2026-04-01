namespace ExcelLibrary;

/// <summary>
/// Specifies the number format category of a cell in Excel.
/// </summary>
public enum NumberFormat
{
    /// <summary>General format with no specific number formatting.</summary>
    General,

    /// <summary>Numeric format for displaying numbers.</summary>
    Number,

    /// <summary>Currency format with currency symbol.</summary>
    Currency,

    /// <summary>Accounting format with aligned currency symbols and decimal points.</summary>
    Accounting,

    /// <summary>Date format for displaying dates.</summary>
    Date,

    /// <summary>Time format for displaying times.</summary>
    Time,

    /// <summary>Percentage format that multiplies by 100 and adds a percent sign.</summary>
    Percentage,

    /// <summary>Fraction format for displaying values as fractions.</summary>
    Fraction,

    /// <summary>Scientific notation format (e.g., 1.23E+10).</summary>
    Scientific,

    /// <summary>Text format that treats the cell value as text.</summary>
    Text,

    /// <summary>Special format for zip codes, phone numbers, etc.</summary>
    Special,

    /// <summary>Custom user-defined format.</summary>
    Custom,

    /// <summary>Format not recognized by the library.</summary>
    Unsupported
}
