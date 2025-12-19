using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class SubBlock : AbstractFunctionBlock
    {
        public SubBlock(
            bool isPublic,
            bool isDefault,
            NameToken name,
            IEnumerable<Parameter> parameters,
            IEnumerable<ICodeBlock> statements)
            : base(isPublic, isDefault, false, name, parameters, statements) { }

        protected override string keyWord
        {
            get { return "Sub"; }
        }
    }
}
