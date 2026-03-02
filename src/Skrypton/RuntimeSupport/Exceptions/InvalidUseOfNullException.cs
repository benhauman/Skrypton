using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when VBScript Null is passed in where it isn't accepted (eg. to CDbl)
    /// </summary>
    [Serializable]
    public sealed class InvalidUseOfNullException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Invalid use of null";

        [Obsolete("do not use it")] public InvalidUseOfNullException() : this(null, innerException: null) { }
        [Obsolete("do not use it")] public InvalidUseOfNullException(Exception innerException) : this(null, innerException) { }

        public InvalidUseOfNullException(string? additionalInformationIfAny) : this(additionalInformationIfAny, innerException: null) { }
        public InvalidUseOfNullException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 94; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private InvalidUseOfNullException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
