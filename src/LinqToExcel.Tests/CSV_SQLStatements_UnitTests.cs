using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.IO;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class CSV_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        string _fileName;

        [TestFixtureSetUp]
        public void fs()
        {
            _fileName = Path.Combine(Path.GetTempPath(), "spreadsheet.csv");
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
        }

        [Test]
        public void Connection_string_data_source_is_directory_of_csv_file()
        {
            var repo = new ExcelRepository(_fileName);
            var people = from p in repo.Worksheet()
                         select p;

            try { people.GetEnumerator(); }
            catch (OleDbException) { }

            string dataSource = GetDataSource();
            Assert.AreEqual(Path.GetDirectoryName(_fileName), dataSource);
        }

        [Test]
        public void Connection_string_extended_properties_have_csv_settings()
        {
            var repo = new ExcelRepository(_fileName);
            var people = from p in repo.Worksheet()
                         select p;

            try { people.GetEnumerator(); }
            catch (OleDbException) { }

            string extendedProperties = GetExtendedProperties();
            Assert.AreEqual("\"text;HDR=Yes;FMT=Delimited;\"", extendedProperties);
        }

        [Test]
        public void Table_name_is_csv_file_name()
        {
            var repo = new ExcelRepository(_fileName);
            var people = from p in repo.Worksheet()
                         select p;

            try { people.GetEnumerator(); }
            catch (OleDbException) { }

            string tableName = GetTableName(GetSQLStatement());
            Assert.AreEqual(Path.GetFileName(_fileName), tableName);
        }

        private string GetTableName(string sqlStatement)
        {
            string[] words = sqlStatement.Split(" ".ToCharArray());
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i] == "FROM")
                    return words[i + 1];
            }
            return "";
        }
    }
}
