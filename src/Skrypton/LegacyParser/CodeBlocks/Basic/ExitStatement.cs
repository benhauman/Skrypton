using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class ExitStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public ExitStatement(ExitableStatementType statementType, int lineIndex)
        {
            if (!Enum.IsDefined(typeof(ExitableStatementType), statementType))
                throw new ArgumentException("Invalid statementType value specified [" + statementType.ToString() + "]");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex");

            StatementType = statementType;
            LineIndex = lineIndex;
        }

        [DataMember] public ExitableStatementType StatementType { get; private set; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        [DataMember] public int LineIndex { get; private set; }

        [DataContract(Namespace = "http://vbs")]
        public enum ExitableStatementType
        {
            [EnumMember] Do,
            [EnumMember] For,
            [EnumMember] Function,
            [EnumMember] Property,
            [EnumMember] Sub
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            return indenter.Indent + "Exit " + StatementType.ToString();
        }
    }
}
