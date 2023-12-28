using System.Data;
using System.CodeDom;

namespace ExcelExport
{
    internal class CodeConverter : CodeConverterBase
    {
        readonly DataTable m_DataTable;

        public CodeConverter(DataTable dataTable, CodeConvertOptions options)
            : base(options)
        {
            m_DataTable = dataTable;
        }

        public override CodeConvertResult ToCode(string splitToken)
        {
            var code = new CodeCompileUnit();
            Exception? ex = null;
            var tableName = m_DataTable.TableName.Trim();

            try
            {
                var codeName = new CodeNamespace(m_Options.Namespace);
                code.Namespaces.Add(codeName);
                codeName.Imports.AddRange(g_Imports.Select(x => new CodeNamespaceImport(x)).ToArray());

                var classCode = WriteClassCode(tableName);
                foreach (var column in m_DataTable.Columns.Cast<DataColumn>())
                {
                    var header = ExcelExportUtils.ReadHeader(column.ColumnName, splitToken);
                    var member = WriteMemberCode(header.type, header.name);
                    classCode.Members.Add(member);
                }
                codeName.Types.Add(classCode);
            }
            catch (Exception ex2)
            {
                ex = ex2;
            }
            
            return new(tableName, code, ex);
        }
    }
}