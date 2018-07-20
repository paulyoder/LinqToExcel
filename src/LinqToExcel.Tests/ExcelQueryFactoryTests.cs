using System.Linq;
using NUnit.Framework;
using System;
using System.IO;
using System.Data;
using LinqToExcel.Domain;

namespace LinqToExcel.Tests
{
    using LinqToExcel.Query;

    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class ExcelQueryFactoryTests
    {
        private string _excelFileName;
        private string _excelFileWithBuiltinWorksheets;
        private string _excelFileWithNamedRanges;

        [SetUp]
        public void s()
        {
            var excelFilesDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            _excelFileWithBuiltinWorksheets = Path.Combine(excelFilesDirectory, "Companies.xlsx");
            _excelFileWithNamedRanges = Path.Combine(excelFilesDirectory, "NamedRanges.xlsx");
        }

        [Test]
        public void throw_argumentnullexception_when_filename_is_null()
        {
            var repo = new ExcelQueryFactory(new LogManagerFactory());
            Assert.That(() => (from r in repo.Worksheet() select r).First(),
            Throws.TypeOf<ArgumentNullException>());
        }

        [Test]
        public void Constructor_sets_filename()
        {
            var repo = new ExcelQueryFactory(@"C:\spreadsheet.xls", new LogManagerFactory());
            Assert.AreEqual(@"C:\spreadsheet.xls", repo.FileName);
        }

        [Test]
        public void Constructor_defaults_UsePersistentConnection_to_false()
        {
            var repo = new ExcelQueryFactory(new LogManagerFactory());
            Assert.AreEqual(false, repo.UsePersistentConnection);
        }

        [Test]
        public void Constructor_defaults_ReadOnly_to_false()
        {
            var repo = new ExcelQueryFactory(new LogManagerFactory());
            Assert.AreEqual(false, repo.ReadOnly);
        }

        [Test]
        public void GetWorksheetNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory(new LogManagerFactory());
            Assert.That(() => factory.GetWorksheetNames(),
            Throws.TypeOf<NullReferenceException>(), "FileName property is not set");

        }

        [Test]
        public void GetColumnNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory(new LogManagerFactory());
            Assert.That(() => factory.GetColumnNames(""),
            Throws.TypeOf<NullReferenceException>(), "FileName property is not set");
        }

        [Test]
        public void GetWorksheetNames_returns_worksheet_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());

            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "ColumnMappings, IMEX Table, Invalid Cast, More Companies, Null Dates, Range1, Sheet1, TrimSpaces",
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetWorksheetNames_does_not_include_builtin_worksheets()
        {
            var excel = new ExcelQueryFactory(_excelFileWithBuiltinWorksheets, new LogManagerFactory());
            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "AutoFiltered, ColumnMappings, MoreCompanies, NullCells, Paul's Worksheet, Sheet1",
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test] //This test is no longer passing. I believe it has something to do with my computer settings
        public void GetWorksheetNames_does_not_include_named_ranges()
        {
            var excel = new ExcelQueryFactory(_excelFileWithNamedRanges, new LogManagerFactory());
            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "Tabelle1, Tabelle3, WS2",
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetNamedRanges()
        {
            var excel = new ExcelQueryFactory(_excelFileWithNamedRanges, new LogManagerFactory());
            var namedRanges = excel.GetNamedRanges(excel.GetWorksheetNames().First());
            Assert.AreEqual(
                "NameCell",
                string.Join(", ", namedRanges.ToArray()));
        }

        [Test]
        public void GetNamedRangeValue()
        {
            var excel = new ExcelQueryFactory(_excelFileWithNamedRanges, new LogManagerFactory());
            var firstCellValue = excel.NamedRangeNoHeader("Tabelle1", "NameCell").First().First().Value;
            Assert.AreEqual(
                "NameCell",
                firstCellValue);
        }

        [Test]
        public void GetWorksheetNames_does_not_delete_apostrophes_in_middle_of_worksheet_name()
        {
            var excel = new ExcelQueryFactory(_excelFileWithBuiltinWorksheets, new LogManagerFactory());
            var worksheetNames = excel.GetWorksheetNames();
            Assert.IsTrue(worksheetNames.Any(x => x == "Paul's Worksheet"));
        }

        [Test]
        public void GetColumnNames_returns_column_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());

            var columnNames = excel.GetColumnNames("Sheet1");
            Assert.AreEqual(
                "Name, CEO, EmployeeCount, StartDate",
                string.Join(", ", columnNames.ToArray()));
        }

        [Test]
        public void StrictMapping_ClassStrict_throws_StrictMappingException_when_property_is_not_mapped_to_column()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.ClassStrict;
            var companies = (from x in excel.Worksheet<CompanyWithCity>()
                             select x);
            Assert.That(() => companies.ToList(),
            Throws.TypeOf<StrictMappingException>(), "'City' property is not mapped to a column");
        }

        [Test]
        public void StrictMapping_ClassStrict_with_additional_unused_worksheet_columns_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.ClassStrict;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyNullable>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_WorksheetStrict_throws_StrictMappingException_when_column_is_not_mapped_to_property()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.WorksheetStrict;
            var companies = (from x in excel.Worksheet<Company>("Null Dates")
                             select x);
            Assert.That(() => companies.ToList(),
            Throws.TypeOf<StrictMappingException>(), "'City' column is not mapped to a property");

        }

        [Test]
        public void StrictMapping_WorksheetStrict_with_additional_unused_class_properties_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.WorksheetStrict;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyWithCity>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_Both_throws_StrictMappingException_when_property_is_not_mapped_to_column()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.Both;
            var companies = (from x in excel.Worksheet<CompanyWithCity>()
                             select x);
            Assert.That(() => companies.ToList(),
            Throws.TypeOf<StrictMappingException>(), "'City' column is not mapped to a property");
        }

        [Test]
        public void StrictMapping_Both_throws_StrictMappingException_when_column_is_not_mapped_to_property()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.Both;
            var companies = (from x in excel.Worksheet<Company>("Null Dates")
                             select x);
            Assert.That(() => companies.ToList(),
            Throws.TypeOf<StrictMappingException>(), "'City' column is not mapped to a property");
        }

        [Test]
        public void StrictMapping_Both_with_column_mappings_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.Both;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<Company>("More Companies")
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_None_with_additional_worksheet_column_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.None;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<Company>("Null Dates")
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_None_with_additional_class_properties_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.StrictMapping = StrictMappingType.None;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyWithCity>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_Not_Explicitly_Set_with_additional_worksheet_column_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<Company>("Null Dates")
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_Not_Explicitly_Set_with_additional_class_properties_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyWithCity>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void IContainsUnmappedCells_FillsInUnmappedCells()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyNameWithUnmappedCells>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
            Assert.AreEqual("ACME", companies[0].Name);
            Assert.AreEqual(3, companies.First().UnmappedCells.Count);
            Assert.AreEqual("Bugs Bunny", companies.First().UnmappedCells["CEO"].Value);
            Assert.AreEqual(25, companies.First().UnmappedCells["EmployeeCount"].Cast<int>());
            Assert.AreEqual(new DateTime(1918, 11, 11), companies.First().UnmappedCells["StartDate"].Cast<DateTime>());
        }

        [Test]
        public void TrimSpaces_Start_TrimsWhiteSpacesAtTheBeginning()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.TrimSpaces = TrimSpacesType.Start;

            var companies = excel.Worksheet<Company>("TrimSpaces").ToList();

            Assert.AreEqual("White Space In Front", companies[0].Name);
        }

        [Test]
        public void TrimSpaces_End_TrimsWhiteSpacesAtTheEnd()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.TrimSpaces = TrimSpacesType.End;

            var companies = excel.Worksheet<Company>("TrimSpaces").ToList();

            Assert.AreEqual("White Space At End", companies[1].Name);
        }

        [Test]
        public void TrimSpaces_Both_TrimsWhiteSpacesOnBothSides()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.TrimSpaces = TrimSpacesType.Both;

            var companies = excel.Worksheet<Company>("TrimSpaces").ToList();

            Assert.AreEqual("White Space On Both Sides", companies[2].Name);
        }

        [Test]
        public void TrimSpaces_None_DoesntTrimWhitespace()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.TrimSpaces = TrimSpacesType.None;

            var companies = excel.Worksheet<Company>("TrimSpaces").ToList();

            Assert.AreEqual(" White Space On Both Sides ", companies[2].Name);
        }

        [Test]
        public void Null_ConstantExpression_in_where_expression_should_not_throw_exception()
        {
            // The C# 6 compiler made some changes to the expression tree output.
            // This changes are largely transparent to consumers if they adhere
            // to the Liskov Substitution Principle.  This test is to specifically
            // address a bug in which the code didn't allow the derived TypedConstantExpression
            // to be substituted for a ConstantExpression.  This resulted in a downstream InvalidCastException
            // when the LinqToExcel invoked Converter.ChangeType since TypedConstantExpression
            // does not implement IConvertible.
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyNullable>()
                             where c.Name != null
                             select c).ToList();

            Assert.IsNotNull(companies);
        }
    }
}
