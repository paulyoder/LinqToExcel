using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [TestFixture]
    public class ExcelVersionTests
    {
        [Test]
        public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_xls_extensions()
        {
            var repo = new ExcelRepository<Company>("spreadsheet.xls");
            Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        }

        [Test]
        public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_XLS_extensions()
        {
            var repo = new ExcelRepository<Company>("spreadsheet.XLS");
            Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        }

        [Test]
        public void FileType_is_set_to_ExcelVersion_Csv_for_files_with_csv_extensions()
        {
            var repo = new ExcelRepository<Company>("spreadsheet.csv");
            Assert.AreEqual(ExcelVersion.Csv, repo.FileType);
        }

        [Test]
        public void FileType_is_set_to_ExcelVersion_Csv_for_files_with_CSV_extensions()
        {
            var repo = new ExcelRepository<Company>("spreadsheet.CSV");
            Assert.AreEqual(ExcelVersion.Csv, repo.FileType);
        }

        [Test]
        public void FileType_is_set_to_ExcelVersion_PreExcel2007_for_files_with_unrecognized_extensions()
        {
            var repo = new ExcelRepository<Company>("spreadsheet.tdl");
            Assert.AreEqual(ExcelVersion.PreExcel2007, repo.FileType);
        }
    }
}
