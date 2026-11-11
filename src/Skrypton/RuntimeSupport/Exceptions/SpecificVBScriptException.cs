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
            : base(GetMessage(basicErrorDescription, additionalInformationIfAny), innerException)
        {
            _basicDescription = basicErrorDescription;
        }

        private readonly string? _basicDescription;

        /// <summary>
        /// The description of the error without the additional information (eg. "Type mismatch") - Err.Description
        /// </summary>
        public string BasicDescription
        {
            get
            {
                if (_basicDescription != null)
                    return _basicDescription;
                int separator = Message.IndexOf(": ", StringComparison.Ordinal);
                return (separator > 0) ? Message.Substring(0, separator) : Message;
            }
        }

        /// <summary>
        /// What Windows VBScript names when it reports the (unhandled) error to the host - "Object required: 'name'", "Type mismatch: 'CInt'",
        /// "Subscript out of range: '[number: 5]'" - or null if it names nothing (issue #84). Err.Description never has it.
        /// </summary>
        public string? HostTarget { get; internal set; }

        /// <summary>
        /// The description of the error as Windows VBScript reports it to the host (cscript, MSScriptControl's Error.Description)
        /// </summary>
        public string HostDescription => (HostTarget == null) ? BasicDescription : BasicDescription + ": '" + HostTarget + "'";

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
