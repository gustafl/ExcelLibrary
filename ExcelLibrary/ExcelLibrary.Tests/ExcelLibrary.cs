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
        private const string FILE = @"C:\Users\Gustaf\Desktop\test2.xlsx";

        private Workbook workbook = null;
        private Workbook workbookWithIncludeHidden = null;

        [TestInitialize]
        [TestCategory("Sheet")]
        public void Initialize()
        {
            this.workbook = new Workbook();
            this.workbook.Open(FILE);

            WorkbookOptions options = new WorkbookOptions();
            options.IncludeHidden = true;
            this.workbookWithIncludeHidden = new Workbook();
            this.workbookWithIncludeHidden.Open(FILE, options);
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
            Sheet sheet = this.workbook.Sheet("Sheet3");
            Assert.AreEqual("Sheet3", sheet.Name);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetId()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            Assert.AreEqual("rId3", sheet.Id);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetSheetPath()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            Assert.AreEqual("xl/worksheets/sheet3.xml", sheet.Path);
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
            Sheet sheet = this.workbook.Sheet("Sheet3");
            Assert.AreSame(this.workbook, sheet.Workbook);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void OpenSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Assert.IsInstanceOfType(sheet, typeof(Sheet));
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowsExcludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
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
            Sheet sheet = this.workbook.Sheet("Sheet3");
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
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
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
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
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
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            string text = row.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnByIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(2);
            string text = column.Cell(2).Value;
            Assert.AreEqual("Banana", text);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCellsExcludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(3, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCellsIncludingHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(5, cells.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByRowAndColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Cell cell = sheet.Cell(2, 2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetCellByName()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Cell cell = sheet.Cell("B2");
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Sheet - Integrity")]
        public void SameNumberOfCellsInRowsAndColumns()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            int rowCellCount = sheet.Rows.SelectMany(r => r.Cells).Count();
            int columnCellCount = sheet.Columns.SelectMany(c => c.Cells).Count();
            Assert.AreEqual(rowCellCount, columnCellCount);
        }

        // TODO: Implement "Sheet - Failure" tests here

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            Assert.AreEqual(2, row.Index);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(8);
            Assert.AreEqual(true, row.Hidden);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            Assert.AreSame(sheet, row.Sheet);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellsInRow()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            int cellCount = row.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellByColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnIndex()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(2);
            Assert.AreEqual(2, column.Index);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnHidden()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(5);
            Assert.AreEqual(true, column.Hidden);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnSheet()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(2);
            Assert.AreSame(sheet, column.Sheet);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellsInColumn()
        {
            Sheet sheet = this.workbookWithIncludeHidden.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(3);
            int cellCount = column.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellByRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        // TODO: Implement test for Cell.Type here.

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellValue()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreEqual("Banana", cell.Value);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellRow()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Row row = sheet.Row(2);
            Cell cell = row.Cell(2);
            Assert.AreSame(row, cell.Row);
        }

        [TestMethod]
        [TestCategory("Cell")]
        public void GetCellColumn()
        {
            Sheet sheet = this.workbook.Sheet("Sheet3");
            sheet.Open();
            Column column = sheet.Column(2);
            Cell cell = column.Cell(2);
            Assert.AreSame(column, cell.Column);
        }
    }
}
