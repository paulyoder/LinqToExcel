using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class MiscTests
    {
        [Test]
        [ExpectedArgumentNullException]
        public void throw_argumentnullexception_when_filename_is_null()
        {
            var repo = new ExcelQueryFactory();
            var first = (from r in repo.Worksheet() select r).First();
        }
    }
}
