using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This is used when an invalid type of parameter is specified (such as a non-positive startIndex for the INSTR function)
    /// </summary>
    [Serializable]
    public sealed class InvalidProcedureCallOrArgumentException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Invalid procedure call or argument";

        [Obsolete("do not use it")]internal InvalidProcedureCallOrArgumentException() : this(null, innerException: null) { }

        public InvalidProcedureCallOrArgumentException(string message) : this(message, innerException: null) { }
        public InvalidProcedureCallOrArgumentException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 5; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private InvalidProcedureCallOrArgumentException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
