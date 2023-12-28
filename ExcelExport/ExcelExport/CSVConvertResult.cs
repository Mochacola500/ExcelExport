
namespace ExcelExport
{
    public class CSVConvertResult
    {
        public bool IsSuccess => Exception == null;
        public string CSV { get; init; }
        public Exception? Exception { get; init; }

        public CSVConvertResult(string csv, Exception? exception)
        {
            CSV = csv;
            Exception = exception;
        }
    }
}
