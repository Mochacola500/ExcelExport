using System.Data;
using System.CodeDom;
using System.Reflection;

namespace ExcelExport
{
    internal class CodeConverter
    {
        static readonly string[] g_Imports = new[]
        {
            "System",
            "System.Collections.Generic",
        };

        readonly DataTable m_DataTable;
        readonly CodeConvertOptions m_Options;

        public CodeConverter(DataTable dataTable, CodeConvertOptions options)
        {
            m_DataTable = dataTable;
            m_Options = options;
        }

        public CodeConvertResult ToCode(string splitToken)
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

        CodeTypeDeclaration WriteClassCode(string name)
        {
            return new()
            {
                TypeAttributes = TypeAttributes.Public,
                IsPartial = true,
                IsClass = true,
                Name = name,
            };
        }

        CodeMemberField WriteMemberCode(string type, string name)
        {
            return new(type, name)
            {
                Attributes = MemberAttributes.Public,
            };
        }

        /*
        CodeMemberMethod WriteMemberMethod(string name, string bodyExpression, params (string type, string name)[] parameters)
        {
            var method = new CodeMemberMethod()
            {
                Attributes = MemberAttributes.Public,
                Name = name,
            };
            method.Statements.Add(new CodeSnippetExpression(bodyExpression));
            method.Parameters.AddRange(parameters.Select(x => new CodeParameterDeclarationExpression(x.type, x.name)).ToArray());
            return method;
        }
        */
    }
}