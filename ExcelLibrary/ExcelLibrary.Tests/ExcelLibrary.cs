using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class ExcelLibrary
    {
        private const string FILE = @"C:\Users\Gustaf\Desktop\test2.xlsx";

        private Workbook workbook = null;
        private Workbook workbookWithOptions = null;

        [TestInitialize]
        [TestCategory("Workbook")]
        public void Initialize()
        {
            this.workbook = new Workbook();
            this.workbook.Open(FILE);

            WorkbookOptions options = new WorkbookOptions();
            options.IncludeHidden = true;
            this.workbookWithOptions = new Workbook();
            this.workbookWithOptions.Open(FILE, options);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void OpenWorkbook()
        {
            Assert.IsInstanceOfType(this.workbook, typeof(Workbook));
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void OpenWorkbookWithOptions()
        {
            Assert.IsInstanceOfType(this.workbookWithOptions, typeof(Workbook));
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void CountSheetsExcludingHidden()
        {
            var sheets = this.workbook.Sheets;
            Assert.AreEqual(3, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void CountSheetsIncludingHidden()
        {
            var sheets = this.workbookWithOptions.Sheets;
            Assert.AreEqual(4, sheets.Count());
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookFile()
        {
            Assert.AreEqual(FILE, this.workbook.File);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookOptionsWithOptionsLoaded()
        {
            WorkbookOptions options = this.workbookWithOptions.Options;
            Assert.AreEqual(true, options.IncludeHidden);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetWorkbookOptionsWithNoOptionsLoaded()
        {
            WorkbookOptions options = this.workbook.Options;
            Assert.AreEqual(false, options.IncludeHidden);
        }

        [TestMethod]
        [TestCategory("Workbook")]
        public void GetSharedStringsCount()
        {
            Assert.AreEqual(1, this.workbook.SharedStrings.Count);
        }
    }
}
