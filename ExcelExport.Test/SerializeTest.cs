
namespace ExcelExport.Test
{
    [TestFixture]
    public class SerializeTest
    {
        ExcelExport m_SimpleData;
        Dictionary<string, string> m_JsonData;

        [SetUp]
        public void Init()
        {
            m_SimpleData = Load("SimpleData.xlsx");
            using var jsonData = Load("JsonData.xlsx");
            m_JsonData = jsonData.ToJson().ToDictionary(x => x.Name, x => x.Json);
        }

        ExcelExport Load(string path)
        {
            var fullPath = Path.Combine("../../../Data", path);
            return ExcelExport.Load(fullPath);
        }

        [TearDown]
        public void CleanUp()
        {
            m_SimpleData.Dispose();
        }

        [Test]
        public void CSVTest1()
        {
            var csvResult = m_SimpleData.ToCSV().First();
            string expected = "a,b,c,d,e\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5\r\n1,2,3,4,5";
            Assert.That(csvResult.CSV, Is.EqualTo(expected));
        }

        [Test]
        public void JsonTest1()
        {
            var jsonResult = m_SimpleData.ToJson().First();
            string expected = "[{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"}]";
            Assert.That(jsonResult.Json, Is.EqualTo(expected));
        }

        [Test]
        public void MemberJsonTest()
        {
            var expected = "{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"}";
            Assert.That(m_JsonData["MemberData"], Is.EqualTo(expected));  
        }

        [Test]
        public void ListJsonTest()
        {
            var expected = "[{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},{\"a\":\"2\",\"b\":\"3\",\"c\":\"4\",\"d\":\"5\",\"e\":\"6\"}]";
            Assert.That(m_JsonData["ListData"], Is.EqualTo(expected));
        }

        [Test]
        public void DictionaryJsonTest()
        {
            var expected = "\"1\":{\"id\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":\"4\",\"e\":\"5\"},\"2\":{\"id\":\"2\",\"b\":\"3\",\"c\":\"4\",\"d\":\"5\",\"e\":\"6\"}";
            Assert.That(m_JsonData["DictionaryData"], Is.EqualTo(expected));
        }
    }
}