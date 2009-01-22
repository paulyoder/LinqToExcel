using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using LinqToExcel.Extensions.Object;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestsOn(typeof(ObjectExtensions))]
    [TestFixture]
    public class ObjectExtensionsTests
    {
        [Test]
        public void As_casts_object_to_generic_argument_type()
        {
            int number = 25;
            Assert.AreEqual(number.As<double>().GetType(), typeof(double));
        }

        [Test]
        [ExpectedException(typeof(InvalidCastException))]
        public void As_throws_exception_on_invalid_cast()
        {
            string nothing = null;
            nothing.As<DateTime>();
        }
    }
}
