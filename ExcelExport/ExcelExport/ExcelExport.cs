using System.Data;
using System.Text;
using ExcelDataReader;

namespace ExcelExport
{
    public class ExcelExport : IDisposable
    {

        static readonly ExcelReaderConfiguration g_ReaderConfig = new()
        {
            FallbackEncoding = Encoding.UTF8
        };

        static readonly ExcelDataSetConfiguration g_DataSetConfig = new()
        {
            FilterSheet = IsCommentSheet,
            ConfigureDataTable = (reader) => g_DataTableConfig
        };

        static readonly ExcelDataTableConfiguration g_DataTableConfig = new()
        {
            FilterColumn = IsComment,
            UseHeaderRow = true,
        };

        readonly DataSet m_DataSet;
        readonly ExcelExportOptions m_Options;

        ExcelExport(DataSet dataSet, ExcelExportOptions options)
        {
            m_DataSet = dataSet;
            m_Options = options;
        }

        public static ExcelExport Load(string path, ExcelExportOptions? option = null)
        {
            using var stream = File.OpenRead(path);
            if (option == null)
            {
                option = new();
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var dataReader = ExcelReaderFactory.CreateReader(stream, g_ReaderConfig);
            var dataTable = dataReader.AsDataSet(g_DataSetConfig);
            return new ExcelExport(dataTable, option);
        }

        public IEnumerable<CSVConvertResult> ToCSV(CSVConvertOptions? options = null)
        {
            if (options == null)
            {
                options = new();
            }
            return m_DataSet.Tables
                .Cast<DataTable>()
                .Select(x => new CSVConverter(x, options).ToCSV(m_Options.SplitToken));
        }

        public IEnumerable<JsonConvertResult> ToJson(JsonConvertOptions? options = null)
        {
            if (options == null)
            {
                options = new();
            }
            return m_DataSet.Tables
                .Cast<DataTable>()
                .Select(x => new JsonConverter(x, options).ToJson(m_Options.SplitToken));
        }

        public IEnumerable<CodeConvertResult> ToCode(CodeConvertOptions? options = null)
        {
            if (options == null)
            {
                options = new();
            }
            return m_DataSet.Tables
                .Cast<DataTable>()
                .Select(x => new CodeConverter(x, options).ToCode(m_Options.SplitToken));
        }

        public void Dispose()
        {
            if (m_DataSet != null)
            {
                m_DataSet.Dispose();
            }
        }

        static bool IsCommentSheet(IExcelDataReader reader, int _)
        {
            var name = reader.Name;
            return !(string.IsNullOrEmpty(name) || name[0] == '#');
        }

        static bool IsComment(IExcelDataReader reader, int index)
        {
            var name = reader.GetString(index);
            return !(string.IsNullOrEmpty(name) || name[0] == '#');
        }
    }
}