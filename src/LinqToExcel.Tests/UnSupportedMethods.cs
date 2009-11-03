using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class UnSupportedMethods
    {
        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Contains() method")]
        public void contains()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Contains(new Company());
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the DefaultIfEmpty() method")]
        public void default_if_empty()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).DefaultIfEmpty().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Distinct() method")]
        public void distinct()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Distinct().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Except() method")]
        public void except()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Except(new List<Company>()).ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Group() method")]
        public void group()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            group c by c.CEO into g
                            select g;
            try { companies.ToList(); }
            catch (OleDbException) { }
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Intersect() method")]
        public void intersect()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c.CEO).Intersect(
                             from d in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select d.CEO)
                             .ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the OfType() method")]
        public void of_type()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).OfType<object>().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Single() method. Use the First() method instead")]
        public void single()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Single();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Union() method")]
        public void union()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Union(
                             from d in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select d)
                             .ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel does not provide support for the Join() method")]
        public void join()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             join d in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             on c.CEO equals d.CEO
                             select d)
                             .ToList();
        }
    }
}
