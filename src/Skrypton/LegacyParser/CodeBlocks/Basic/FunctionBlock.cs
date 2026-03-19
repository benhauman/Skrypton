using Skrypton.LegacyParser.Tokens.Basic;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class FunctionBlock : AbstractFunctionBlock
    {
        public FunctionBlock(
            bool isPublic,
            bool isDefault,
            NameToken name,
            IEnumerable<Parameter> parameters,
            IEnumerable<ICodeBlock> statements)
            : base(isPublic, isDefault, true, name, parameters, statements) { }

        protected override string keyWord
        {
            get { return "Function"; }
        }
    }
}
