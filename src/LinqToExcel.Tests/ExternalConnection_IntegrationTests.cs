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
	public class ExternalConnection_IntegrationTests
	{
		private IExcelQueryFactory _factory;
		private OleDbConnection _externalConnection;

		[TestFixtureSetUp]
		public void fs()
		{
			string testDirectory = AppDomain.CurrentDomain.BaseDirectory;
			string excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
			string excelFileName = Path.Combine(excelFilesDirectory, "Companies.xlsm");

			//create a connection to be shared across all queries
			_externalConnection = new OleDbConnection(
				string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
					excelFileName));

			_factory = new ExcelQueryFactory(excelFileName)
				{
					ExternalConnection = _externalConnection
				};
		}

		[Test]
		public void WorksheetRangeNoHeader_returns_7_companies()
		{
			var companies = from c in _factory.WorksheetRangeNoHeader("A1", "D8", "Sheet1")
			                select c;

			Assert.AreEqual(7, companies.Count());
		}

		[Test]
		public void WorksheetRangeNoHeader_can_query_sheet_500_times_on_same_connection()
		{
			IQueryable<Row> rows = null;

			for (int i = 0; i < 500; i++)
			{
				rows = from cm in _factory.WorksheetRange("C1", "I8")
				       select cm;

				if (rows.Count() != 7)
				{
					Assert.AreEqual(7, rows.Count());	
				}
			}

			Assert.AreEqual(_externalConnection.GetHashCode(), _factory.ExternalConnection.GetHashCode());
		}

		[TestFixtureTearDown]
		public void td()
		{
			//dispose of the connection
			if (_factory.ExternalConnection != null)
			{
				_factory.ExternalConnection.Dispose();
			}
		}
	}
}
