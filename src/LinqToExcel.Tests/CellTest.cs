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
    public class CellTest
    {
        [Test]
        public void Constructor_sets_cell_value()
        {
            Cell newCell = new Cell("hello");
            Assert.AreEqual("hello", newCell.Value);
        }

        [Test]
        public void ValueAs_converts_cell_value_type_to_generic_argument_type()
        {
            Cell newCell = new Cell("2");
            Assert.AreEqual(2, newCell.ValueAs<int>());
            Assert.AreEqual(typeof(int), newCell.ValueAs<int>().GetType());
        }

        [Test]
        public void ValueAs_returns_default_generic_value_when_value_is_null()
        {
            Cell newCell = new Cell(null);
            Assert.AreEqual(0, newCell.ValueAs<int>());
        }

        [Test]
        public void ValueAs_returns_default_generic_value_when_value_is_DBNull()
        {
            Cell newCell = new Cell(DBNull.Value);
            Assert.AreEqual(0, newCell.ValueAs<int>());
        }
    }
}
