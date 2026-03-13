using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    internal sealed class ConstStatement : ICodeBlock
    {
        public ConstStatement(IEnumerable<ConstValueInitialisation> values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            Values = values.ToList().AsReadOnly();
            if (!Values.Any())
                throw new ArgumentException("Empty values set - invalid");
            if (Values.Any(v => v == null))
                throw new ArgumentException("Null reference encountered in values set");
        }

        /// <summary>
        /// This will never be null, empty nor contain any nulls
        /// </summary>
        [DataMember] internal IEnumerable<ConstValueInitialisation> Values { get; private set; }

        [DataContract(Namespace = "http://vbs")]
        internal sealed class ConstValueInitialisation
        {
            public ConstValueInitialisation(NameToken name, IToken value)
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));

                if (!(value is DateLiteralToken) && !(value is NumericValueToken) && !(value is StringToken))
                {
                    var builtInValueToken = value as BuiltInValueToken;
                    if ((builtInValueToken == null) || !builtInValueToken.IsAcceptableAsConstValue)
                        throw new ArgumentException($"Invalid CONST value - must be a literal or supported builtin value type and not '{value.GetType().Name}'. Line:{value.LineIndex}:{value.Content}", nameof(value));
                }

                Name = name ?? throw new ArgumentNullException(nameof(name));
                Value = value;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            [DataMember] public NameToken Name { get; private set; }

            /// <summary>
            /// This will never be null, it will always be a literal value - a boolean, number, string or date - or one of the acceptable builtin values, such as Empty
            /// or Null (acceptables values have a true IsAcceptableAsConstValue property on BuiltInValueToken instance)
            /// </summary>
            [DataMember] public IToken Value { get; private set; }

            public override string ToString()
            {
                return base.ToString() + ":" + Name;
            }
        }

        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(IBaseSourceGenerationContext generationContext)
        {
            if (generationContext == null) throw new ArgumentNullException(nameof(generationContext));
            return string.Format(CultureInfo.InvariantCulture,
                "{0}Const {1}",
                generationContext.Indent,
                string.Join(", ", Values.Select(v => v.Name.Content + " = " + TokenValueAsVbsCode(v.Value)))
            );
        }

        private static string TokenValueAsVbsCode(IToken token)
        {
            if (token is StringToken stringToken)
            {
                return StringExtensions.ToLiteral(stringToken.Content);
            }
            else if (token is NumericValueToken numericToken)
            {
                return numericToken.Content;
            }
            else if (token is DateLiteralToken dateToken)
            {
                return dateToken.Content;
            }
            else if (token is BuiltInValueToken builtInValueToken)
            {
                return builtInValueToken.Content;
            }
            throw new NotSupportedException($"{token.GetType().Name} : {token.Content}");
        }
    }
}
