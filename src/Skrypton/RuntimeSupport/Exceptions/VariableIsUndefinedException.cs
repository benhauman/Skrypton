using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// Error 500, "Variable is undefined: 'name'" - a variable that is not declared in a script with Option Explicit (with VBScript's rule for
    /// undeclared variables, see ScriptControlConfiguration.UndeclaredVariablesAreErrors and issue #86). ScriptLine is the 1-based line.
    /// </summary>
    [Serializable]
#pragma warning disable CA1032 // the variable and the line are required
    public sealed class VariableIsUndefinedException : SpecificVBScriptException
#pragma warning restore CA1032
    {
        private const string BASIC_ERROR_DESCRIPTION = "Variable is undefined";

        public VariableIsUndefinedException(string variableName, int scriptLine)
            : base(BASIC_ERROR_DESCRIPTION, $"'{variableName}' (line {scriptLine})", innerException: null)
        {
            HostTarget = variableName;
            ScriptLine = scriptLine;
        }

        public override int ErrorNumber { get { return 500; } }

        public int ScriptLine { get; }

        private VariableIsUndefinedException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
