using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelLibrary.Tests
{
    [TestClass]
    public class NumberFormats
    {
        private const string FILE = @"..\..\Input\test2.xlsx";

        private Workbook workbook = null;
        private Sheet sheet = null;
        private Column column = null;

        [TestInitialize]
        public void Initialize()
        {
            this.workbook = new Workbook();
            this.workbook.Open(FILE);
            this.sheet = this.workbook.Sheet("Sheet1");
            this.column = sheet.Column(2);
        }

        [TestMethod]
        [TestCategory("NumberFormats")]
        public void General()
        {
            Cell cell = this.column.Cell(1);
            string val = cell.Value.Replace(".", ",");
            decimal number = decimal.Parse(val);
            Assert.AreEqual(123.45m, number);
        }

        [TestMethod]
        [TestCategory("NumberFormats")]
        public void Number()
        {
            Cell cell = this.column.Cell(2);
            string val = cell.Value.Replace(".", ",");
            decimal number = decimal.Parse(val);
            Assert.AreEqual(123.45m, number);
        }

        [TestMethod]
        [TestCategory("NumberFormats")]
        public void Currency()
        {
            Cell cell = this.column.Cell(3);
            string val = cell.Value.Replace(".", ",");
            decimal number = decimal.Parse(val);
            Assert.AreEqual(123.45m, number);
        }

        [TestMethod]
        [TestCategory("NumberFormats")]
        public void Accounting()
        {
            Cell cell = this.column.Cell(4);
            string val = cell.Value.Replace(".", ",");
            decimal number = decimal.Parse(val);
            Assert.AreEqual(123.45m, number);
        }
    }
}
