using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class DefaultOptions
    {
        private const string FILE = @"..\..\Input\test1.xlsx";

        private Workbook workbook = null;

        [TestInitialize]
        public void Initialize()
        {
            this.workbook = new Workbook();
            this.workbook.Open(FILE);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void OpenWorkbook()
        {
            Assert.IsInstanceOfType(this.workbook, typeof(Workbook));
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookFile()
        {
            Assert.AreEqual(FILE, this.workbook.File);
        }



        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSheets()
        {
            var sheets = this.workbook.Sheets;
            Assert.AreEqual(3, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSharedStrings()
        {
            Assert.AreEqual(35, this.workbook.SharedStrings.Count);
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
        public void GetRows()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            var rows = sheet.Rows;
            foreach (Row row in rows)
                Debug.WriteLine("{0}", row.Index);
            Assert.AreEqual(4, rows.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumns()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            var columns = sheet.Columns;
            foreach (Column column in columns)
                Debug.WriteLine("{0}", column.Index);
            Assert.AreEqual(3, columns.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            string text = row.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(2);
            string text = column.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCells()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(4, cells.Count());
        }


        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByRowAndColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell(2, 2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByName()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell("B2");
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void SameNumberOfCellsInRowsAndColumns()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            int rowCellCount = sheet.Rows.SelectMany(r => r.Cells).Count();
            int columnCellCount = sheet.Columns.SelectMany(c => c.Cells).Count();
            Assert.AreEqual(rowCellCount, columnCellCount);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingRowByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(100);
            Assert.IsNull(row);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingColumnByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(100);
            Assert.IsNull(column);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetHiddenRowWithIncludeHiddenFalse()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(8);
            Assert.IsNull(row);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetHiddenColumnWithIncludeHiddenFalse()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(5);
            Assert.IsNull(column);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellsOnEmptySheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(0, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingCellByRowIndexAndColumnIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell(100, 100);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetNonExistingCellByName()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell("A100");
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellInHiddenRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell(8, 3);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void TryGetCellInHiddenColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Cell cell = sheet.Cell(2, 5);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Assert.AreEqual(2, row.Index);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Assert.AreSame(sheet, row.Sheet);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellsInRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            int cellCount = row.Cells.Count();
            Assert.AreEqual(1, cellCount);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellByColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void TryGetRowCellInHiddenColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Cell cell = row.Cell(5);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(2);
            Assert.AreEqual(2, column.Index);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(2);
            Assert.AreSame(sheet, column.Sheet);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellsInColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(3);
            int cellCount = column.Cells.Count();
            Assert.AreEqual(1, cellCount);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellByRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void TryGetColumnCellInHiddenRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(3);
            Cell cell = column.Cell(8);
            Assert.IsNull(cell);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellValue()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreSame(row, cell.Row);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreSame(column, cell.Column);
        }

        // TODO: GetCellFormat()
        // TODO: GetCellValueFromSharedStrings()
        // TODO: GetCellValueFromInlineString()
    }
}
