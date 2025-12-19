using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class FunctionBlock : AbstractFunctionBlock
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
