using System.Text;

namespace ExcelExport
{
    internal class CsvSerializer
    {
        readonly char m_Seperator;
        readonly char[] m_QuoteTable;
        readonly StringBuilder m_CsvBuilder;

        CsvSerializer(char seperator, char[] quoteTable)
        {
            m_Seperator = seperator;
            m_QuoteTable = quoteTable;
            m_CsvBuilder = new StringBuilder(1024 * 1024 * 256);
        }

        public static string Serialize(char seperator, string[] columnNames, IEnumerable<string[]> dataArray)
        {
            var quoteTable = new char[]
            {
                seperator,
                '\r',
                '\n',
                '\"',
            };
            return new CsvSerializer(seperator, quoteTable).Serialize(columnNames, dataArray);
        }

        string Serialize(string[] columnNames, IEnumerable<string[]> dataArray)
        {
            if (columnNames.Length == 0)
            {
                return "";
            }
            m_CsvBuilder.Clear();

            // Write Header.
            m_CsvBuilder.Append(columnNames[0]);
            for (int i = 1; i < columnNames.Length; ++i)
            {
                m_CsvBuilder.Append(m_Seperator);
                m_CsvBuilder.Append(columnNames[i]);
            }
            m_CsvBuilder.AppendLine();

            // Write Data.
            if (dataArray.Any())
            {
                var firstItem = dataArray.First();
                WriteData(firstItem);
                var e = dataArray.GetEnumerator();
                // Skip first array.
                e.MoveNext();
                while (e.MoveNext())
                {
                    m_CsvBuilder.AppendLine();
                    WriteData(e.Current);
                }
            }
            return m_CsvBuilder.ToString();
        }

        void WriteData(string[] columns)
        {
            if (columns.Length == 0)
            {
                return;
            }
            WriteField(columns[0]);
            for (int i = 1; i < columns.Length; ++i)
            {
                m_CsvBuilder.Append(m_Seperator);
                var column = columns[i];
                WriteField(column);
            }
        }

        void WriteField(string column)
        {
            bool needQuote = column.IndexOfAny(m_QuoteTable) != -1;
            if (needQuote)
            {
                int lastIndex = m_CsvBuilder.Length;
                m_CsvBuilder.Append('"')
                    .Append(column)
                    .Append('"');

                if (column.Contains("\""))
                {
                    m_CsvBuilder.Replace("\"", "\"\"", lastIndex + 1, column.Length);
                }
            }
            else
            {
                m_CsvBuilder.Append(column);
            }
        }
    }
}