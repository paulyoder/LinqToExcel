using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [TestCategory("Unit")]
    [TestFixture]
    public class RowTest
    {
        IDictionary<string, int> _columnMappings;
        IList<Cell> _cells;

        [SetUp]
        public void s()
        {
            _columnMappings = new Dictionary<string, int>();
            _columnMappings["Name"] = 0;
            _columnMappings["Favorite Sport"] = 1;
            _columnMappings["Age"] = 2;

            _cells = new List<Cell>();
            _cells.Add(new Cell("Paul"));
            _cells.Add(new Cell("Ultimate Frisbee"));
            _cells.Add(new Cell(25));
        }

        [Test]
        public void index_maps_cells_correctly()
        {
            var row = new Row(_cells, _columnMappings);
            Assert.AreEqual(_cells[0], row[0]);
            Assert.AreEqual(_cells[1], row[1]);
            Assert.AreEqual(_cells[2], row[2]);
        }

        [Test]
        public void column_name_index_maps_cells_correctly()
        {
            var row = new Row(_cells, _columnMappings);
            Assert.AreEqual(_cells[0], row["Name"]);
            Assert.AreEqual(_cells[1], row["Favorite Sport"]);
            Assert.AreEqual(_cells[2], row["Age"]);
        }

        [Test]
        [ExpectedArgumentException("Column name does not exist: First Name")]
        public void invalid_column_name_index_throws_argument_exception()
        {
            var newRow = new Row(_cells, _columnMappings);
            var temp = newRow["First Name"];
        }
    }
}
