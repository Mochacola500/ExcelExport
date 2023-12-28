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
            var tableName = m_DataTable.TableName.Trim();

            try
            {
                var columnNames = m_DataTable.Columns
                    .Cast<DataColumn>()
                    .Select(x => ExcelExportUtils.ReadHeader(x.ColumnName, splitToken).name)
                    .ToArray();

                var dataArray = m_DataTable.Rows
                    .Cast<DataRow>()
                    .Select(x => x.ItemArray.Select(x => x?.ToString() ?? " ").ToArray());

                var type = (columnNames.Length > 0 && columnNames[0] == m_Options.IdToken) ?
                    ExportCollectionType.Dictionary :
                    dataArray.Count() == 1 ? 
                    ExportCollectionType.Member :
                    ExportCollectionType.List;

                if (m_Options.ExportToArray)
                {
                    type = ExportCollectionType.List;
                }

                var sr = new JsonSerializer(type, columnNames);
                json = sr.Serialize(dataArray);
            }
            catch (Exception ex2)
            {
                ex = ex2;
            }

            return new(tableName, json, ex);
        }
    }
}