using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This will be raised when a FOR EACH target can not be enumerated over
    /// </summary>
    [Serializable]
    public sealed class ObjectNotCollectionException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object not a collection";

        [Obsolete("do not use it")] private ObjectNotCollectionException() : this(null, innerException: null) { }
        public ObjectNotCollectionException(string message) : this(message, innerException: null) { }
        public ObjectNotCollectionException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 451; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm
        private ObjectNotCollectionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
