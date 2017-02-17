using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using MbUnit.Framework;

namespace LinqToExcel.Tests
{
	[Author("Andrew Corkery", "andrew.corkery@gmail.com")]
	[FixtureCategory("Integration")]
	[TestFixture]
	public class PersistentConnection_IntegrationTests
	{
		private IExcelQueryFactory _factory;

		[TestFixtureSetUp]
		public void fs()
		{
			string testDirectory = AppDomain.CurrentDomain.BaseDirectory;
			string excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
			string excelFileName = Path.Combine(excelFilesDirectory, "Companies.xlsm");

			_factory = new ExcelQueryFactory(excelFileName, new LogManagerFactory());
            _factory.UsePersistentConnection = true;
		}

		[Test]
		public void WorksheetRangeNoHeader_returns_7_companies()
		{
			var companies = from c in _factory.WorksheetRangeNoHeader("A2", "D8", "Sheet1")
			                select c;

			Assert.AreEqual(7, companies.Count());
		}

		[Test]
		public void WorksheetRangeNoHeader_can_query_sheet_500_times_on_same_connection()
		{
			IQueryable<Row> rows = null;

			int totalRows = 0;

			for (int i = 0; i < 500; i++)
			{
				rows = from cm in _factory.WorksheetRange("A2", "D8")
				       select cm;

				totalRows += rows.Count();
			}

			Assert.AreEqual((500*7), totalRows);
		}

		[TestFixtureTearDown]
		public void td()
		{
			//dispose of the factory (and persistent connection)
			_factory.Dispose();
		}
	}
}
