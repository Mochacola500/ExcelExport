
namespace ExcelExport
{
    internal static class ExcelExportUtils
    {
        public static (string type, string name) ReadHeader(string text, string splitToken)
        {
            var i = text.IndexOf(splitToken);
            if (i == -1)
            {
                throw new Exception($"Split token missing.\ntext:{text}");
            }
            var type = text.Substring(i).Trim();
            var name = text.Substring(0, i).Trim();
            return (type, name);
        }
    }
}