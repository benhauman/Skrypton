using System;
using System.Linq;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This is used when the source code included Err.Raise - those errors are translated into exceptions of this type
    /// </summary>
    [Serializable]
    public sealed class CustomException : SpecificVBScriptException
    {
        private const string DEFAULT_SOURCE = "(null)";
        private const string DEFAULT_DESCRIPTION = "Unknown runtime error";

        private readonly int _errorNumber;

#pragma warning disable IDE0051 // CA1032 wants the standard exception constructors - they are private so that they are not used
        [Obsolete("do not use it")] private CustomException() { }
        [Obsolete("do not use it")] private CustomException(string message) : base(message) { }
        [Obsolete("do not use it")] private CustomException(string message, Exception innerException) : base(message, innerException) { }
#pragma warning restore IDE0051
        public CustomException(int number, string source, string description) : base(GetMessage(source, description), additionalInformationIfAny: null, innerException: null)
        {
            if (number == 0)
                throw new ArgumentOutOfRangeException(nameof(number));

            _errorNumber = number;
            ErrSource = source ?? "";
            ErrDescription = description ?? "";
        }

        /// <summary>
        /// An error of a built-in component whose source and description are as given, also when they are empty (Err.Raise without them gets those
        /// of the VBScript error, or "Unknown runtime error") - eg. WshShell.Run of a missing program (issue #165)
        /// </summary>
        internal static CustomException AsGiven(int number, string source, string description) => new CustomException(number, source, description) { IsAsGiven = true };

        /// <summary>
        /// The source and the description are as given, also when they are empty (see AsGiven)
        /// </summary>
        internal bool IsAsGiven { get; private set; }

        /// <summary>
        /// The source given to Err.Raise ("" if there was none) - Err.Source
        /// </summary>
        public string ErrSource { get; } = "";

        /// <summary>
        /// The description given to Err.Raise ("" if there was none) - Err.Description
        /// </summary>
        public string ErrDescription { get; } = "";

        /// <summary>
        /// This will never be zero (but it may be a negative or positive value)
        /// </summary>
        public override int ErrorNumber { get { return _errorNumber; } }

        private CustomException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        private static string GetMessage(string source, string description)
        {
            if (string.IsNullOrWhiteSpace(source))
                source = DEFAULT_SOURCE;
            if (string.IsNullOrWhiteSpace(description))
                description = DEFAULT_DESCRIPTION;

            return string.Join(": ", new[] { source, description }.Where(v => !string.IsNullOrWhiteSpace(v)));
        }
    }
}
