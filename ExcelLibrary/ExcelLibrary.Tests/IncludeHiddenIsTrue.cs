using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class IncludeHiddenIsTrue
    {
        private const string FILE = @"..\..\Input\test1.xlsx";

        private Workbook workbook = null;
        private WorkbookOptions options = null;

        [TestInitialize]
        public void Initialize()
        {
            options = new WorkbookOptions();
            options.IncludeHidden = true;
            options.LoadSheets = true;
            this.workbook = new Workbook();
            this.workbook.Open(FILE, options);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSheetsIncludingHidden()
        {
            var sheets = this.workbook.Sheets;
            Assert.AreEqual(4, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetRowsIncludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            var rows = sheet.Rows;
            foreach (Row row in rows)
            {
                Debug.WriteLine("{0}", row.Index);
            } 
            Assert.AreEqual(5, rows.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetColumnsIncludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            var columns = sheet.Columns;
            foreach (Column column in columns)
            {
                Debug.WriteLine("{0}", column.Index);
            }  
            Assert.AreEqual(4, columns.Count());
        }

        [TestMethod]
        [TestCategory("Sheet")]
        public void GetAllCellsIncludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            IEnumerable<Cell> cells = sheet.Cells;
            Assert.AreEqual(6, cells.Count());
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetCellsInRowIncludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(2);
            int cellCount = row.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Row")]
        public void GetRowHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Row row = sheet.Row(8);
            Assert.AreEqual(true, row.Hidden);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetCellsInColumnIncludingHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(3);
            int cellCount = column.Cells.Count();
            Assert.AreEqual(2, cellCount);
        }

        [TestMethod]
        [TestCategory("Column")]
        public void GetColumnHidden()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            Column column = sheet.Column(5);
            Assert.AreEqual(true, column.Hidden);
        }
    }
}
