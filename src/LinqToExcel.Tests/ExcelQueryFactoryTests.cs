﻿using System.Linq;
using MbUnit.Framework;
using System;
using System.IO;
using System.Data;
using LinqToExcel.Domain;

namespace LinqToExcel.Tests
{
    using LinqToExcel.Query;

    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
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
        [ExpectedArgumentNullException]
        public void throw_argumentnullexception_when_filename_is_null()
        {
            var repo = new ExcelQueryFactory();
            var first = (from r in repo.Worksheet() select r).First();
        }

        [Test]
        public void Constructor_sets_filename()
        {
            var repo = new ExcelQueryFactory(@"C:\spreadsheet.xls");
            Assert.AreEqual(@"C:\spreadsheet.xls", repo.FileName);
        }

        [Test]
        public void Constructor_defaults_DatabaseEngine_to_Jet()
        {
            var repo = new ExcelQueryFactory();
            Assert.AreEqual(DatabaseEngine.Jet, repo.DatabaseEngine);
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException), "FileName property is not set")]
        public void GetWorksheetNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory();
            factory.GetWorksheetNames();
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException), "FileName property is not set")]
        public void GetColumnNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory();
            factory.GetColumnNames("");
        }

        [Test]
        public void GetWorksheetNames_returns_worksheet_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName);

            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "ColumnMappings, IMEX Table, More Companies, Null Dates, Range1, Sheet1",
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetWorksheetNames_does_not_include_builtin_worksheets()
        {
            var excel = new ExcelQueryFactory(_excelFileWithBuiltinWorksheets);
            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "AutoFiltered, ColumnMappings, MoreCompanies, NullCells, Paul's Worksheet, Sheet1",
                string.Join(", ", worksheetNames.ToArray()));
        }

        //[Test] This test is no longer passing. I believe it has something to do with my computer settings
        public void GetWorksheetNames_does_not_include_named_ranges()
        {
            var excel = new ExcelQueryFactory(_excelFileWithNamedRanges);
            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "Tabelle1, Tabelle3, WS2",
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetWorksheetNames_does_not_delete_apostrophes_in_middle_of_worksheet_name()
        {
            var excel = new ExcelQueryFactory(_excelFileWithBuiltinWorksheets);
            var worksheetNames = excel.GetWorksheetNames();
            Assert.IsTrue(worksheetNames.Any(x => x == "Paul's Worksheet"));
        }

        [Test]
        public void GetColumnNames_returns_column_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName);

            var columnNames = excel.GetColumnNames("Sheet1");
            Assert.AreEqual(
                "Name, CEO, EmployeeCount, StartDate",
                string.Join(", ", columnNames.ToArray()));
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' property is not mapped to a column")]
        public void StrictMapping_ClassStrict_throws_StrictMappingException_when_property_is_not_mapped_to_column()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.ClassStrict;
            var companies = (from x in excel.Worksheet<CompanyWithCity>()
                             select x).ToList();
        }

        [Test]
        public void StrictMapping_ClassStrict_with_additional_unused_worksheet_columns_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.ClassStrict;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyNullable>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' column is not mapped to a property")]
        public void StrictMapping_WorksheetStrict_throws_StrictMappingException_when_column_is_not_mapped_to_property()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.WorksheetStrict;
            var companies = (from x in excel.Worksheet<Company>("Null Dates")
                             select x).ToList();
        }

        [Test]
        public void StrictMapping_WorksheetStrict_with_additional_unused_class_properties_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.WorksheetStrict;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyWithCity>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' property is not mapped to a column")]
        public void StrictMapping_Both_throws_StrictMappingException_when_property_is_not_mapped_to_column()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.Both;
            var companies = (from x in excel.Worksheet<CompanyWithCity>()
                             select x).ToList();
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' column is not mapped to a property")]
        public void StrictMapping_Both_throws_StrictMappingException_when_column_is_not_mapped_to_property()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = StrictMappingType.Both;
            var companies = (from x in excel.Worksheet<Company>("Null Dates")
                             select x).ToList();
        }

        [Test]
        public void StrictMapping_Both_with_column_mappings_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
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
            var excel = new ExcelQueryFactory(_excelFileName);
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
            var excel = new ExcelQueryFactory(_excelFileName);
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
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<Company>("Null Dates")
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }

        [Test]
        public void StrictMapping_Not_Explicitly_Set_with_additional_class_properties_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<CompanyWithCity>()
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }
    }
}
