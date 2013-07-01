﻿using System;
using System.Collections.Generic;
using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
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

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]")]
        public void distinct_on_whole_row()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Distinct().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]")]
        public void distinct_on_no_header()
        {
            var excel = new ExcelQueryFactory("");
            var companies = (from c in excel.WorksheetNoHeader()
                             select c).Distinct().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]")]
        public void distinct_on_no_header_with_selected_column()
        {
            var excel = new ExcelQueryFactory("");
            var companies = (from c in excel.WorksheetNoHeader()
                             select c[0]).Distinct().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]")]
        public void distinct_on_row()
        {
            var excel = new ExcelQueryFactory("");
            var companies = (from c in excel.Worksheet()
                             select c).Distinct().ToList();
        }

        [Test]
        [ExpectedException(typeof(NotSupportedException), "LinqToExcel only provides support for the Distinct() method when it's mapped to a class and a single property is selected. [e.g. (from row in excel.Worksheet<Person>() select row.FirstName).Distinct()]")]
        public void distinct_on_row_with_selected_column()
        {
            var excel = new ExcelQueryFactory("");
            var companies = (from c in excel.Worksheet()
                             select c["Name"]).Distinct().ToList();
        }
    }
}
