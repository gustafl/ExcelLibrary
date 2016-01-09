using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class LoadSheetsIsFalse
    {
        private const string FILE = @"..\..\Input\test1.xlsx";

        private Workbook workbook = null;
        private WorkbookOptions options = null;

        [TestInitialize]
        public void Initialize()
        {
            options = new WorkbookOptions();
            options.IncludeHidden = false;
            options.LoadSheets = false;
            this.workbook = new Workbook();
            this.workbook.Open(FILE, options);
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
        public void TryGetRowsBeforeCallingSheetOpen()
        {
            Sheet sheet = this.workbook.Sheet("Sheet1");
            IEnumerable<Row> rows = sheet.Rows;
            Assert.AreEqual(0, rows.Count());
        }
    }
}
