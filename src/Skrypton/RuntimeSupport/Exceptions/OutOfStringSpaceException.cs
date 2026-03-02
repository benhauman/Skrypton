using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when a string is required that would be too long (whether that comes from concatenating other strings or by using the STRING method)
    /// </summary>
    [Serializable]
    public sealed class OutOfStringSpaceException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Out of string space";

        [Obsolete("do not use it")] private OutOfStringSpaceException() : this(null, innerException: null) { }
        public OutOfStringSpaceException(string message) : this(message, innerException: null) { }
        public OutOfStringSpaceException(string additionalInformationIfAny, Exception innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 14; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private OutOfStringSpaceException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
