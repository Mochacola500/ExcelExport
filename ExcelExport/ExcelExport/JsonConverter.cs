using System.Data;

namespace ExcelExport
{
    internal class JsonConverter
    {
        readonly DataTable m_DataTable;
        readonly JsonConvertOptions m_Options;

        public JsonConverter(DataTable dataTable, JsonConvertOptions options)
        {
            m_DataTable = dataTable;
            m_Options = options;
        }

        public JsonConvertResult ToJson(string splitToken)
        {
            string json = "";
            Exception? ex = null;
            try
            {
                var columnNames = m_DataTable.Columns
                    .Cast<DataColumn>()
                    .Select(x => ExcelExportUtils.ReadHeader(x.ColumnName, splitToken).name)
                    .ToArray();

                var dataArray = m_DataTable.Rows
                    .Cast<DataRow>()
                    .Select(x => x.ItemArray.Cast<string>().ToArray());

                var sr = new JsonSerializer(columnNames);
                json = sr.Serialize(dataArray);
            }
            catch (Exception ex2)
            {
                ex = ex2;
            }

            return new(json, ex);
        }
    }
}