using System.Text;

namespace ExcelExport
{
    internal class JsonSerializer
    {
        readonly string[] m_ColumnNames;
        readonly StringBuilder m_JsonBuilder;

        public JsonSerializer(string[] columnNames)
        {
            m_ColumnNames = columnNames;
            m_JsonBuilder = new StringBuilder(1024 * 1024 * 256);
        }

        public string Serialize(IEnumerable<string[]> dataArray)
        {
            m_JsonBuilder.Clear();
            m_JsonBuilder.Append("[");
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
            m_JsonBuilder.Append("]");
            return m_JsonBuilder.ToString();
        }

        void WriteData(string[] columns)
        {
            if (columns.Length == 0)
            {
                return;
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
            m_JsonBuilder.Append('\"')
                .Append(m_ColumnNames[i])
                .Append("\":")
                .Append('\"')
                .Append(field)
                .Append('\"');
        }
    }
}