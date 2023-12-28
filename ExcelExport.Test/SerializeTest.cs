
namespace ExcelExport.Test
{
    [TestFixture]
    public class SerializeTest
    {
        ExcelExport m_Export;

        [SetUp]
        public void Setup()
        {
            m_Export = ExcelExport.Load("../../../SimpleData.xlsx");
        }

        [Test]
        public void CSVTest1()
        {
            var csvResult = m_Export.ToCSV().First();
            string expected = "a,b,c,d,e\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5";
            Assert.That(csvResult.CSV, Is.EqualTo(expected));
        }

        [Test]
        public void JsonTest1()
        {
            var jsonResult = m_Export.ToJson().First();
            string expected = "[{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"}]";
            Assert.That(jsonResult.Json, Is.EqualTo(expected));
        }
    }
}