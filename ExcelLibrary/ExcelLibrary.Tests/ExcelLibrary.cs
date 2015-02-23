using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class ExcelLibrary
    {
        private const string FILE = @"..\..\Input\test1.xlsx";
        private const string FILE_ALLTYPES = @"..\..\Input\testAllTypes.xlsx";

        private Workbook workbook = null;
        private Workbook workbookWithIncludeHidden = null;

        private Workbook workbook_AllTypes = null;
        private Workbook workbookWithIncludeHidden_AllTypes = null;

        [TestInitialize]
        [TestCategory("Sheet")]
        public void Initialize()
        {
            // regular test file initialize
            this.workbook = new Workbook();
            this.workbook.Open(FILE);

            WorkbookOptions options = new WorkbookOptions();
            options.IncludeHidden = true;
            this.workbookWithIncludeHidden = new Workbook();
            this.workbookWithIncludeHidden.Open(FILE, options);

            // all types test file initialize
            this.workbook_AllTypes = new Workbook();
            this.workbook_AllTypes.Open(FILE_ALLTYPES);

            WorkbookOptions options_AllTypes = new WorkbookOptions();
            options_AllTypes.IncludeHidden = true;
            this.workbookWithIncludeHidden_AllTypes = new Workbook();
            this.workbookWithIncludeHidden_AllTypes.Open(FILE_ALLTYPES, options_AllTypes);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void OpenWorkbook()
        {
            Assert.IsInstanceOfType(this.workbook, typeof(Workbook));
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void OpenWorkbookWithIncludeHidden()
        {
            Assert.IsInstanceOfType(this.workbookWithIncludeHidden, typeof(Workbook));
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookFile()
        {
            Assert.AreEqual(FILE, this.workbook.File);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookOptionsIncludeHidden()
        {
            WorkbookOptions options = this.workbookWithIncludeHidden.Options;
            Assert.AreEqual(true, options.IncludeHidden);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSheetsExcludingHidden()
        {
            var sheets = this.workbook.Sheets;
            Assert.AreEqual(3, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSheetsIncludingHidden()
        {
            var sheets = this.workbookWithIncludeHidden.Sheets;
            Assert.AreEqual(4, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSharedStrings()
        {
            Assert.AreEqual(5, this.workbook.SharedStrings.Count);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetName()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Assert.AreEqual("Sheet1", sheet.Name);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetId()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Assert.AreEqual("rId1", sheet.Id);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetPath()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Assert.AreEqual("xl/worksheets/sheet1.xml", sheet.Path.ToLower());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet4");
            Assert.AreEqual(true, sheet.Hidden);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetWorkbook()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Assert.AreSame(this.workbook, sheet.Workbook);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void OpenSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Assert.IsInstanceOfType(sheet, typeof(Sheet));
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowsExcludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            var rows = sheet.Rows;
            foreach (Row row in rows)
                Debug.WriteLine("{0}", row.Index);
            Assert.AreEqual(3, rows.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnsExcludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            var columns = sheet.Columns;
            foreach (Column column in columns)
                Debug.WriteLine("{0}", column.Index);
            Assert.AreEqual(3, columns.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowsIncludingHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            var rows = sheet.Rows;
            foreach (Row row in rows)
                Debug.WriteLine("{0}", row.Index);
            Assert.AreEqual(4, rows.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnsIncludingHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            var columns = sheet.Columns;
            foreach (Column column in columns)
                Debug.WriteLine("{0}", column.Index);
            Assert.AreEqual(4, columns.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            string text = row.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(2);
            string text = column.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCellsExcludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(3, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCellsIncludingHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(5, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByRowAndColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell(2, 2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByName()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell("B2");
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void SameNumberOfCellsInRowsAndColumns()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            int rowCellCount = sheet.Rows.SelectMany(r => r.Cells).Count();
            int columnCellCount = sheet.Columns.SelectMany(c => c.Cells).Count();
            Assert.AreEqual(rowCellCount, columnCellCount);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetRowsBeforeCallingSheetOpen()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            IEnumerable<Row> rows = sheet.Rows;
            Assert.AreEqual(0, rows.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingRowByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(1);
            Assert.IsNull(row);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingColumnByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(1);
            Assert.IsNull(column);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetHiddenRowWithIncludeHiddenFalse()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(8);
            Assert.IsNull(row);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetHiddenColumnWithIncludeHiddenFalse()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(5);
            Assert.IsNull(column);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellsOnEmptySheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet2");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(0, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingCellByRowIndexAndColumnIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell(1, 1);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingCellByName()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell("A1");
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellInHiddenRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell(8, 3);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellInHiddenColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Cell cell = sheet.Cell(2, 5);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Assert.AreEqual(2, row.Index);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(8);
            Assert.AreEqual(true, row.Hidden);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Assert.AreSame(sheet, row.Sheet);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellsInRow()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            int cellCount = row.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellByColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void TryGetRowCellInHiddenColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(5);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(2);
            Assert.AreEqual(2, column.Index);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(5);
            Assert.AreEqual(true, column.Hidden);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(2);
            Assert.AreSame(sheet, column.Sheet);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellsInColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(3);
            int cellCount = column.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellByRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void TryGetColumnCellInHiddenRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(3);
            Cell cell = column.Cell(8);
            Assert.IsNull(cell);
        }

        // TODO: Implement test for Cell.Type here.

        // Cell.Type tests
        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellType()
        {
            Sheet sheet = this.workbook_AllTypes.Sheet("Sheet1");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            foreach (Cell c in cells)
            {
                Assert.IsFalse(c.Type.Equals(DBNull.Value));
                Assert.IsFalse(c.Type.Equals("m"));
            }
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void CheckCellTypes()
        {
            Sheet sheet = this.workbook_AllTypes.Sheet("Sheet1");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            foreach (Cell c in cells)
            {
                if (c.Row.Index == 1)
                {
                    if (c.Column.Index == 1)
                        Assert.AreSame("General", c.Type);
                    if (c.Column.Index == 2)
                        Assert.AreSame("General", c.Type);
                    if (c.Column.Index == 3)
                        Assert.AreSame("Number", c.Type);
                    if (c.Column.Index == 4)
                        Assert.AreSame("Currency", c.Type);
                    if (c.Column.Index == 5)
                        Assert.AreSame("Accounting", c.Type);
                    if (c.Column.Index == 6)
                        Assert.AreSame("Date", c.Type);
                    if (c.Column.Index == 7)
                        Assert.AreSame("Time", c.Type);
                    if (c.Column.Index == 8)
                        Assert.AreSame("Percentage", c.Type);
                    if (c.Column.Index == 9)
                        Assert.AreSame("Fraction", c.Type);
                    if (c.Column.Index == 10)
                        Assert.AreSame("Scientific", c.Type);
                    if (c.Column.Index == 11)
                        Assert.AreSame("Text", c.Type);
                }
            }
        }


        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellValue()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreSame(row, cell.Row);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            sheet.Open();
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreSame(column, cell.Column);
        }
    }
}
