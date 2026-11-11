using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Skrypton.CSharpWriter;
using Skrypton.LegacyParser.CodeBlocks.Basic;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// IScriptControl.Procedures - the Subs and Functions declared in the code added with AddCode, like MSScriptControl's Procedures collection
    /// (see issue #47). Items can be looked up by name (case-insensitive) or by 1-based index.
    /// </summary>
    internal sealed class ScriptControlProcedures : Procedures
    {
        private readonly Procedure[] _procedures;

        private ScriptControlProcedures(Procedure[] procedures)
        {
            _procedures = procedures ?? throw new ArgumentNullException(nameof(procedures));
        }

        internal static ScriptControlProcedures FromScript(CultureInfo culture, string scriptContent)
        {
            if (culture == null)
                throw new ArgumentNullException(nameof(culture));
            if (string.IsNullOrWhiteSpace(scriptContent))
                return new ScriptControlProcedures(Array.Empty<Procedure>());

            return new ScriptControlProcedures(
                DefaultTranslator.Parse(culture, scriptContent)
                    .OfType<AbstractFunctionBlock>()
                    .Where(block => !(block is PropertyBlock))
                    .Select(block => (Procedure)new ScriptControlProcedure(block.Name.Content, block.Parameters.Count(), block.HasReturnValue))
                    .ToArray()
            );
        }

        public int Count => _procedures.Length;

        public Procedure this[object Index]
        {
            get
            {
                if (Index is string name)
                {
                    Procedure? procedure = _procedures.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                    return procedure ?? throw new KeyNotFoundException($"The procedure '{name}' was not found (there are {_procedures.Length} procedure(s): {string.Join(", ", _procedures.Select(p => p.Name))})");
                }
                if ((Index is int) || (Index is short) || (Index is long) || (Index is byte))
                {
                    long oneBasedIndex = Convert.ToInt64(Index, CultureInfo.InvariantCulture);
                    if ((oneBasedIndex < 1) || (oneBasedIndex > _procedures.Length))
                        throw new ArgumentOutOfRangeException(nameof(Index), $"The procedure index {oneBasedIndex} is out of range (1 to {_procedures.Length})");
                    return _procedures[oneBasedIndex - 1];
                }
                throw new ArgumentException($"A procedure must be looked up by name or by 1-based index, not by a value of type '{Index?.GetType().FullName ?? "null"}'", nameof(Index));
            }
        }

        public IEnumerator<Procedure> GetEnumerator() => ((IEnumerable<Procedure>)_procedures).GetEnumerator();

        IEnumerator IScriptProcedureCollection.GetEnumerator() => _procedures.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _procedures.GetEnumerator();

        private sealed class ScriptControlProcedure : Procedure
        {
            public ScriptControlProcedure(string name, int numArgs, bool hasReturnValue)
            {
                Name = name;
                NumArgs = numArgs;
                HasReturnValue = hasReturnValue;
            }

            public string Name { get; }
            public int NumArgs { get; }
            public bool HasReturnValue { get; }
        }
    }
}
