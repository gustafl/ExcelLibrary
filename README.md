# ExcelLibrary

A lightweight, dependency-free .NET library for reading Excel workbooks (.xlsx).

```csharp
// One-liner with static factory
using var workbook = Workbook.Open("data.xlsx");
var value = workbook.Sheet("Sheet1")?.Cell("B2")?.Value;
```

## Installation

```shell
dotnet add package ExcelLibrary
```

Or via the NuGet Package Manager:

```shell
Install-Package ExcelLibrary
```

## Quick Start

### Reading cells

```csharp
using var workbook = Workbook.Open("Book1.xlsx");

// Access by sheet name and cell address
var sheet = workbook.Sheet("Sheet1");
var cell = sheet?.Cell("A1");
Console.WriteLine(cell?.Value);

// Or by row and column index (1-based)
var value = sheet?.Cell(2, 3)?.Value;
```

### Iterating rows and cells

```csharp
using var workbook = Workbook.Open("report.xlsx");

foreach (var sheet in workbook.Sheets)
{
    Console.WriteLine($"Sheet: {sheet.Name}");

    foreach (var row in sheet.Rows)
    {
        foreach (var cell in row.Cells)
        {
            Console.Write($"{cell.Value}\t");
        }
        Console.WriteLine();
    }
}
```

### Including hidden elements

By default, hidden sheets, rows, and columns are excluded. Use `WorkbookOptions` to include them:

```csharp
using var workbook = Workbook.Open("data.xlsx", new WorkbookOptions { IncludeHidden = true });

// Now hidden sheets, rows, and columns are accessible
var hiddenSheet = workbook.Sheet("HiddenSheet");
```

### Lazy loading sheets

For large workbooks, you can defer loading sheet data until needed:

```csharp
using var workbook = Workbook.Open("large-file.xlsx", new WorkbookOptions { LoadSheets = false });

// Sheet metadata is available, but rows/cells are not yet loaded
var sheet = workbook.Sheet("Sheet1");

// Load the sheet data when needed
sheet?.Open();
```

## Features

- **Zero dependencies** — Uses only built-in .NET APIs
- **LINQ-friendly** — Collections like `Sheets`, `Rows`, and `Cells` are `IEnumerable<T>`
- **Visibility-aware** — Respects hidden sheets, rows, and columns by default
- **Lazy loading** — Optionally defer sheet loading for better performance
- **Well-tested** — Comprehensive test suite with 70+ unit tests
- **Fully documented** — XML documentation for IntelliSense support

## API Reference

| Class | Description |
|-------|-------------|
| `Workbook` | Represents an Excel file; provides access to sheets and metadata |
| `Sheet` | A worksheet containing rows, columns, and cells |
| `Row` | A row with access to its cells |
| `Column` | A column with access to its cells |
| `Cell` | A single cell with its value and format |
| `WorkbookOptions` | Configuration for loading workbooks |

## Limitations

This library focuses on **reading** Excel files. The following are out of scope:

- File formats other than `.xlsx`
- Writing/modifying workbooks
- Formula evaluation
- Cell formatting/styles

## Requirements

- .NET 8.0 or later

## License

MIT
