
namespace ExcelExport
{
    public class JsonConvertResult
    {
        public bool IsSuccess => Exception == null;
        public string Json { get; init; }
        public Exception? Exception { get; init; }

        public JsonConvertResult(string json, Exception? exception)
        {
            Json = json;
            Exception = exception;
        }
    }
}