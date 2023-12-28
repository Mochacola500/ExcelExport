
namespace ExcelExport
{
    public class JsonConvertResult
    {
        public bool IsSuccess => Exception == null;
        public string Name { get; init; }
        public string Json { get; init; }
        public Exception? Exception { get; init; }

        public JsonConvertResult(string name, string json, Exception? exception)
        {
            Name = name;
            Json = json;
            Exception = exception;
        }
    }
}