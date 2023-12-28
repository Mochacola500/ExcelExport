using System.Data;

namespace ExcelExport
{
    internal class CSVConverter
    {
        readonly DataTable m_DataTable;
        readonly CSVConvertOptions m_Options;

        public CSVConverter(DataTable dataTable, CSVConvertOptions options)
        {
            m_DataTable = dataTable;
            m_Options = options;    
        }

        public CSVConvertResult ToCSV(string splitToken)
        {
            string csv = "";
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
                    .Select(x => x.ItemArray.Select(x => x?.ToString() ?? "").ToArray());

                csv = CsvSerializer.Serialize(m_Options.Seperator, columnNames, dataArray);
            }
            catch (Exception ex2)
            {
                ex = ex2;
            }

            return new(tableName, csv, ex);
        }
    }
}