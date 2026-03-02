using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when Nothing is passed in where a VBScript-value-type reference is expected
    /// </summary>
    [Serializable]
    public sealed class ObjectVariableNotSetException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object variable not set";

        [Obsolete("do not use it")] private ObjectVariableNotSetException() : this(null, innerException: null) { }
        public ObjectVariableNotSetException(string message) : this(message, innerException: null) { }
        public ObjectVariableNotSetException(string additionalInformationIfAny, Exception innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 91; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private ObjectVariableNotSetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
