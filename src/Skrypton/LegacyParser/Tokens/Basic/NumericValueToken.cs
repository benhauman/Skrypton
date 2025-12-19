using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public sealed class NumericValueToken : AtomToken
    {
        /// <summary>
        /// The constructor must take the original string content representing the number since it's important to differentiate between "1" and "1.0"
        /// (where the first is declared as an "Integer" in VBScript and the latter as a "Double")
        /// </summary>
        public NumericValueToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
        public NumericValueToken(StringUpper contentUpper, int lineIndex) : base(contentUpper.Original.Trim().ToUpperX(), WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");

            double numericValue;
            if (!double.TryParse(contentUpper.Original, out numericValue))
                throw new ArgumentException("content must be a string representation of a numeric value");

            //Value = numericValue;
        }

        public static int CompareNumericValueToken(NumericValueToken x, NumericValueToken y)
        {
            var base_cmp = CompareAtomTokens(x, y);
            if (base_cmp == 0)
            {
                if (x.Value == y.Value)
                    return 0;
                return 8;
            }
            return base_cmp;
        }

        /// <summary>
        /// This will never be null or blank, nor have any leading or trailing whitespace. It will always be parseable as a numeric value.
        /// </summary>
        /// public new string Content { get { return base.Content; } }

        private double? numericValue;
        public double Value
        {
            get
            {
                if (!numericValue.HasValue)
                {
                    numericValue = double.Parse(this.Content);
                }
                return numericValue.Value;
            }
        }
    }
}