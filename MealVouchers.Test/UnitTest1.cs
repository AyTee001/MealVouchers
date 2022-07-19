namespace MealVouchers.Test
{
    [TestClass]
    public class HelpingMethodsTest
    {
        [TestMethod]
        public void RealMonthNameTest1()
        {
            string[] monthNames = new string[] { "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember" };

            var month = "März";
            var expected = "Februar";

            var resultMonth = HelpingMethods.RealMonthName(monthNames, month);

            Assert.AreEqual(expected, resultMonth);
        }
        [TestMethod]
        public void RealMonthNameTest2()
        {

            string[] monthNames = new string[] { "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember" };

            var month = "abc";

            Assert.ThrowsException<InvalidDataException>(() => HelpingMethods.RealMonthName(monthNames, month));
        }
    }

    [TestClass]
    public class ExcelExtract
    {
        [TestMethod]
        public void GetStringsTest() 
        {
            Console.WriteLine(Environment.CurrentDirectory);
            var docName = @"..\..\..\Essensmarken_AT.xlsx";
            var sheet = "Januar";
            string[][] aimCells = new string[][] { new string[] { "F2", "F23" }, new string[] { "G2", "G23" } };
            List<string> expectedFor1 = new() {"75", "111"};
            List<string> expectedFor2 = new() {"7575", "777"};

            var excelExractor = new ExcelExtractor(docName, sheet);
            var prints  = excelExractor.GetStrings(aimCells);
            var item1 = prints[0]; 
            var item2 = prints[1];

            Assert.IsTrue(item1.SequenceEqual(expectedFor1));
            Assert.IsTrue(item2.SequenceEqual(expectedFor2));
        }
    }
}
