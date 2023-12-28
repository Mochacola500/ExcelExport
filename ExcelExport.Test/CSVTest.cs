
namespace ExcelExport.Test
{
    [TestFixture]
    public class CSVTest
    {
        ExcelExport m_Export;

        [SetUp]
        public void Setup()
        {
            m_Export = ExcelExport.Load("../../../SimpleData.xlsx");
        }

        [Test]
        public void Test1()
        {
            var csvResult = m_Export.ToCSV().First();
            string expected = "a,b,c,d,e\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5";
            Assert.That(csvResult.CSV, Is.EqualTo(expected));
        }
    }
}