using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This occurs when a non-Object reference is provided where an Object reference is required (eg. with the "IS" comparison)
    /// </summary>
    [Serializable]
    public sealed class ObjectRequiredException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object required";

        [Obsolete("do not use it")] private ObjectRequiredException() : this(null, innerException: null) { }
        public ObjectRequiredException(string message) : this(message, innerException: null) { }
        public ObjectRequiredException(string additionalInformationIfAny, Exception innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 424; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private ObjectRequiredException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
