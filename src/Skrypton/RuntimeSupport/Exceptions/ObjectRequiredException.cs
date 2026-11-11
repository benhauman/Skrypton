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
        public ObjectRequiredException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        /// <summary>
        /// Error 424 with what VBScript names when it reports the error to the host (HostTarget, verified with cscript): the name of a local
        /// variable or parameter, the value of a global variable ('5', '' for Empty), null for Nothing / Null in a global variable - Err.Description
        /// in the script is "Object required" either way
        /// </summary>
        public static ObjectRequiredException ForScriptTarget(string? scriptTarget, string? additionalInformationIfAny)
        {
            return new ObjectRequiredException(additionalInformationIfAny, innerException: null) { HostTarget = scriptTarget };
        }

        public override int ErrorNumber { get { return 424; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private ObjectRequiredException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
