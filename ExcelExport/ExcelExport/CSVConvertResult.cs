
namespace ExcelExport
{
    public class CSVConvertResult
    {
        public bool IsSuccess => Exception == null;
        public string Name { get; init; }
        public string CSV { get; init; }
        public Exception? Exception { get; init; }

        public CSVConvertResult(string name, string csv, Exception? exception)
        {
            Name = name;
            CSV = csv;
            Exception = exception;
        }
    }
}
