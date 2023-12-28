using System.Text;

namespace ExcelExport
{
    internal class JsonSerializer
    {
        readonly ExportCollectionType m_Type;
        readonly string[] m_ColumnNames;
        readonly StringBuilder m_JsonBuilder;

        public JsonSerializer(ExportCollectionType type, string[] columnNames)
        {
            m_Type = type;
            m_ColumnNames = columnNames;
            m_JsonBuilder = new StringBuilder(1024 * 1024 * 256);
        }

        public string Serialize(IEnumerable<string[]> dataArray)
        {
            m_JsonBuilder.Clear();
            if (m_Type == ExportCollectionType.List)
            {
                m_JsonBuilder.Append("[");
            }
            if (dataArray.Any())
            {
                var firstItem = dataArray.First();
                WriteData(firstItem);
                var e = dataArray.GetEnumerator();
                // Skip first array.
                e.MoveNext();
                while (e.MoveNext())
                {
                    m_JsonBuilder.Append(',');
                    WriteData(e.Current);
                }
            }
            if (m_Type == ExportCollectionType.List)
            {
                m_JsonBuilder.Append("]");
            }
            return m_JsonBuilder.ToString();
        }

        void WriteData(string[] columns)
        {
            if (columns.Length == 0)
            {
                return;
            }
            if (m_Type == ExportCollectionType.Dictionary)
            {
                m_JsonBuilder.Append('\"')
                .Append(columns[0])
                .Append("\":");
            }
            m_JsonBuilder.Append("{");
            WriteField(columns[0], 0);
            for (int i = 1; i < columns.Length; ++i)
            {
                var column = columns[i];
                m_JsonBuilder.Append(',');
                WriteField(column, i);
            }
            m_JsonBuilder.Append('}');
        }

        void WriteField(string field, int i)
        {
            field = field.Replace("\"", "\\\"");
            m_JsonBuilder.Append('\"')
                .Append(m_ColumnNames[i])
                .Append("\":")
                .Append('\"')
                .Append(field)
                .Append('\"');
        }
    }
}