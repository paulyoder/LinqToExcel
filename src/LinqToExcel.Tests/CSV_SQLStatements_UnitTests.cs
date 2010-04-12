using System.Linq;
using MbUnit.Framework;
using System.IO;
using System.Data;

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
            var people = from p in ExcelQueryFactory.Worksheet(null, _fileName, null)
                         select p;

            try { people.GetEnumerator(); }
            catch (DataException) { }

            var dataSource = GetDataSource();
            Assert.AreEqual(Path.GetDirectoryName(_fileName), dataSource);
        }

        [Test]
        public void Connection_string_extended_properties_have_csv_settings()
        {
            var people = from p in ExcelQueryFactory.Worksheet(null, _fileName, null)
                         select p;

            try { people.GetEnumerator(); }
            catch (DataException) { }

            var extendedProperties = GetExtendedProperties();
            Assert.AreEqual("\"text;HDR=YES;FMT=Delimited;IMEX=1\"", extendedProperties);
        }

        [Test]
        public void Table_name_is_csv_file_name()
        {
            var people = from p in ExcelQueryFactory.Worksheet(null, _fileName, null)
                         select p;

            try { people.GetEnumerator(); }
            catch (DataException) { }

            var tableName = GetTableName(GetSQLStatement());
            Assert.AreEqual(Path.GetFileName(_fileName), tableName);
        }

        private string GetTableName(string sqlStatement)
        {
            var words = sqlStatement.Split(" ".ToCharArray());
            for (var i = 0; i < words.Length; i++)
            {
                if (words[i] == "FROM")
                    return words[i + 1]
                        .Replace("[", "")
                        .Replace("]", "");
            }
            return "";
        }
    }
}
