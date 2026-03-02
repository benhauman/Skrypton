using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    [Serializable]
    public sealed class VBScriptDivisionByZeroException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object variable not set";

        [Obsolete("do not use it")] private VBScriptDivisionByZeroException() : this(null, innerException: null) { }
        public VBScriptDivisionByZeroException(string message) : this(message, innerException: null) { }
        public VBScriptDivisionByZeroException(string additionalInformationIfAny, Exception innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 11; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private VBScriptDivisionByZeroException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
