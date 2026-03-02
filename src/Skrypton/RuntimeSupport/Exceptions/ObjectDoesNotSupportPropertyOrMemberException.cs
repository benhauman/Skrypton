using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when a non-value type is passed in where a VBScript-value-type reference is expected, where the reference does not have a default
    /// function or property that can be retrieved without any arguments
    /// </summary>
    [Serializable]
    public sealed class ObjectDoesNotSupportPropertyOrMemberException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object doesn't support this property or method";

        [Obsolete("do not use it.")] private ObjectDoesNotSupportPropertyOrMemberException() : this(null, innerException: null) { }
        public ObjectDoesNotSupportPropertyOrMemberException(string message) : this(message, innerException: null) { }
        public ObjectDoesNotSupportPropertyOrMemberException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 438; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private ObjectDoesNotSupportPropertyOrMemberException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
