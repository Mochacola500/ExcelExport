using System.CodeDom;

namespace ExcelExport
{
    public record CodeConvertResult
    {
        public bool IsSuccess => Exception == null;
        public string Name { get; init; }
        public CodeCompileUnit CodeCompileUnit { get; init; }
        public Exception? Exception { get; init; }

        public CodeConvertResult(string name, CodeCompileUnit codeCompileUnit, Exception? exception)
        {
            Name = name;
            CodeCompileUnit = codeCompileUnit;
            Exception = exception;
        }
    }
}