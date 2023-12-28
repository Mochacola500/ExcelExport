using System.CodeDom;
using System.Reflection;

namespace ExcelExport
{
    internal abstract class CodeConverterBase
    {
        protected static readonly string[] g_Imports = new[]
        {
            "System",
            "System.Collections.Generic",
        };

        protected readonly CodeConvertOptions m_Options;

        protected CodeConverterBase(CodeConvertOptions options)
        {
            m_Options = options;
        }

        public abstract CodeConvertResult ToCode(string splitToken);

        protected CodeTypeDeclaration WriteClassCode(string name)
        {
            return new()
            {
                TypeAttributes = TypeAttributes.Public,
                IsPartial = true,
                IsClass = true,
                Name = name,
            };
        }

        protected CodeMemberField WriteMemberCode(string type, string name)
        {
            return new(type, name)
            {
                Attributes = MemberAttributes.Public,
            };
        }

        protected CodeMemberMethod WriteMemberMethod(string name, string bodyExpression, params (string type, string name)[] parameters)
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
    }
}