using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class UnSupportedMethods
    {
        [Test]
        public void contains()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c).Contains(new Company()),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Contains() method");
        }

        [Test]
        public void default_if_empty()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                             select c).DefaultIfEmpty();
            Assert.That(() => companies.ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the DefaultIfEmpty()");
        }

        [Test]
        public void except()
        {
             Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                                select c).Except(new List<Company>()).ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Group() method");
        }

        [Test]
        public void group()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                            group c by c.CEO into g
                            select g;
            try
            {
                Assert.That(() => companies.ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Group() method");
            }
            catch (OleDbException) { }
        }

        [Test]
        public void intersect()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c.CEO).Intersect(
                             from d in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                             select d.CEO)
                             .ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Intersect() method");
        }

        [Test]
        public void of_type()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c).OfType<object>().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the OfType() method");
        }

        [Test]
        public void single()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c).Single(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Single() method. Use the First() method instead");
        }

        [Test]
        public void union()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c).Union(
                             from d in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                             select d)
                             .ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Union() method");
        }

        [Test]
        public void join()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               join d in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               on c.CEO equals d.CEO
                               select d)
                             .ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel does not provide support for the Join() method");
        }

        [Test]
        public void distinct_on_whole_row()
        {
            Assert.That(() => (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null, new LogManagerFactory())
                               select c).Distinct().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
        }

        [Test]
        public void distinct_on_no_header()
        {
            var excel = new ExcelQueryFactory("", new LogManagerFactory());
            Assert.That(() => (from c in excel.WorksheetNoHeader()
                               select c).Distinct().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
        }

        [Test]
        public void distinct_on_no_header_with_selected_column()
        {
            var excel = new ExcelQueryFactory("", new LogManagerFactory());
            Assert.That(() => (from c in excel.WorksheetNoHeader()
                               select c[0]).Distinct().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
        }

        [Test]
        public void distinct_on_row()
        {
            var excel = new ExcelQueryFactory("", new LogManagerFactory());
            Assert.That(() => (from c in excel.Worksheet()
                               select c).Distinct().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
        }

        [Test]
        public void distinct_on_row_with_selected_column()
        {
            var excel = new ExcelQueryFactory("", new LogManagerFactory());
            Assert.That(() => (from c in excel.Worksheet()
                               select c["Name"]).Distinct().ToList(),
                Throws.TypeOf<NotSupportedException>(), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]");
        }
    }
}
