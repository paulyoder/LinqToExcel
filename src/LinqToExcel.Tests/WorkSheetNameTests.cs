using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;

namespace LinqToExcel.Tests {

   [Author("Paul Yoder", "paulyoder@gmail.com")]
   [Category("Integration")]
   [TestFixture]
   public class WorkSheetNameTests {
      private String _filesDirectory;

      [OneTimeSetUp]
      public void Setup() {
         var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
         _filesDirectory = Path.Combine(testDirectory, "ExcelFiles");
      }

      [Test]
      public void WorkSheetNamesAreDecodedCorrectly() {
         var fileName = Path.Combine(_filesDirectory, "WorksheetNames.xlsx");

         var workbook = new ExcelQueryFactory(fileName, new LogManagerFactory());
         var worksheetNames = workbook.GetWorksheetNames();

         CollectionAssert.AreEqual(
            new [] {
               " ' ",
               "$woot$",
               "Emb$dded",
               "Ends with $'\"",
               "Ends with a $",
               "Has a $ in it"
            },
            worksheetNames
         );
      }
   }
}
