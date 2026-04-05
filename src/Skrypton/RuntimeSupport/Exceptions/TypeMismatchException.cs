using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport
{
    /// <summary>
    /// This occurs when a conversion from one type to another is attempted that fails (eg. passing "a" to CDl)
    /// </summary>
    [Serializable]
    public sealed class TypeMismatchException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Type mismatch";

        [Obsolete("do not use it")] private TypeMismatchException() : this(null, innerException: null) { }
        public TypeMismatchException(string message) : this(message, innerException: null) { }
        public TypeMismatchException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 13; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private TypeMismatchException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
