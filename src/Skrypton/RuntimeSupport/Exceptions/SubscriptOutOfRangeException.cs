using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when an invalid array index is requested (eg. if UBOUND(a, 2) is called with a is a one-dimensional array)
    /// </summary>
    [Serializable]
    public sealed class SubscriptOutOfRangeException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Subscript out of range";

        [Obsolete("do not use it")] private SubscriptOutOfRangeException() : this(null, innerException: null) { }
        public SubscriptOutOfRangeException(string message) : this(message, innerException: null) { }
        public SubscriptOutOfRangeException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 9; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private SubscriptOutOfRangeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
