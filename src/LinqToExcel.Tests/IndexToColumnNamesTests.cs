using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Tests
{
    [TestFixture]
    [Category("Unit")]
    [Author("Alberto Chvaicer")]
    public class IndexToColumnNamesTests
    {
        [Test]
        public void valid_excel_column_indexes()
        {
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(1), "A");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(2), "B");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(3), "C");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(25), "Y");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(26), "Z");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(27), "AA");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(28), "AB");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(51), "AY");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(52), "AZ");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(53), "BA");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(54), "BB");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(701), "ZY");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(702), "ZZ");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(703), "AAA");
                Assert.AreEqual(LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(704), "AAB");
        }

        [Test]
        public void negative_integer_should_throw_exception()
        {
            Assert.That(() => LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(-1), Throws.ArgumentException, "Index should be a positive integer");
        }

        [Test]
        public void zero_should_throw_exception()
        {
            Assert.That(() => LinqToExcel.Query.ExcelUtilities.ColumnIndexToExcelColumnName(-1), Throws.ArgumentException, "Index should be a positive integer");
        }
    }
}
