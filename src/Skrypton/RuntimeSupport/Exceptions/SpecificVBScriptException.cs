using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport
{
    [Serializable]
    public abstract class SpecificVBScriptException : Exception
    {
        protected SpecificVBScriptException() { }
        protected SpecificVBScriptException(string message) : base(message) { }
        protected SpecificVBScriptException(string message, Exception innerException) : base(message, innerException) { }

        protected SpecificVBScriptException(string basicErrorDescription, string? additionalInformationIfAny, Exception? innerException)
            : base(GetMessage(basicErrorDescription, additionalInformationIfAny), innerException) { }

        protected SpecificVBScriptException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public abstract int ErrorNumber { get; }

        private static string GetMessage(string basicErrorDescription, string? additionalInformationIfAny)
        {
            if (basicErrorDescription == null)
                throw new ArgumentNullException(nameof(basicErrorDescription));

            var message = basicErrorDescription;
            if (!string.IsNullOrWhiteSpace(additionalInformationIfAny))
                message += ": " + additionalInformationIfAny!.Trim();
            return message;
        }
    }
}
