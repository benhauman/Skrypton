using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    public sealed class VariableDeclaration
    {
        internal VariableDeclaration(NameToken name, VariableDeclarationScopeOptions scope, IEnumerable<uint>? constantDimensionsIfAny, ConstStatement.ConstValueInitialisation? isConst, IToken? initializationValue)
        {
            if (!Enum.IsDefined(typeof(VariableDeclarationScopeOptions), scope))
                throw new ArgumentOutOfRangeException(nameof(scope));

            if (initializationValue != null)
            {
                if (initializationValue is NumericValueToken numToken || initializationValue is StringToken strToken)
                {
                    // keep
                    InitializationValue = initializationValue;
                }
                else
                {
                    throw new InvalidOperationException($"{name.Content} ({initializationValue.GetType().Name}) = {initializationValue.Content}");
                }

            }
            else
            {
                if (isConst != null)
                {
                    throw new ArgumentNullException(nameof(initializationValue), $"Const '{name.Content}' without initialization value. Line:{name.LineIndex}");
                }
            }


            Name = name ?? throw new ArgumentNullException(nameof(name));
            Scope = scope;
            ConstantDimensionsIfAny = (constantDimensionsIfAny == null) ? null : constantDimensionsIfAny.ToList().AsReadOnly();
            IsConstant = isConst != null;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        public VariableDeclarationScopeOptions Scope { get; private set; }

        /// <summary>
        /// This will be null if this was not an array declaration and may be an empty set if it is an uninitialized array declaration
        /// (array declarations with specified dimensions will always be non-negative integer constants when Dim, Private or Public is
        /// used, otherwise a VBScript compile error will have been raised - ReDim may be used to specify variable dimensions, but they
        /// will be represented by a VariableDeclaration with no dimensions and a separate statement to set the reference to an array)
        /// </summary>
        public IEnumerable<uint>? ConstantDimensionsIfAny { get; private set; }

        public IToken? InitializationValue { get; private set; }

        public bool IsConstant { get; private set; }
    }
}
