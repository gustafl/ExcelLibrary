# ExcelLibrary

This is a small C# library made to simplify reading from and writing to Excel workbooks (.xlsx). Here's an example to get you started:

    Workbook workbook = new Workbook();
    workbook.Open("Book1.xlsx");
    Sheet sheet = workbook.Sheet("Sheet1");
    Row row = sheet.Row(2);
    Cell cell = row.Cell(3);
    string text = cell.Value;

See the wiki for more examples.

## Features

* No dependencies except .NET Framework 4.5. Easy to include in other solutions.
* Built and extendable with LINQ. Most collections in the library (e.g. `Workbook.Sheets` or `Row.Cells`) is of type `IEnumerable<T>`, which  allows you to use LINQ queries to find exactly what you need.
* Respects the visibility of sheets, rows and columns. Set the `IncludeHidden` option to `true`to return hidden objects.
* Well-documented. A software library is only as useful as its documentation.

## Limitations

The following things have been considered outside the scope of the project:

* All file formats except `.xslx`.
* Formulas
* Formatting properties

The following things are planned but not yet implemented features:

* All write functionality (writing to cells, adding new sheets, adding and deleting rows and columns and so on).
* Support for data types (the `Cell.Type` property).

Also, the library will not create new workbooks _per se_, but the same can be achieved by including a template workbook in your project and copy it whenever you need to create a workbook.

