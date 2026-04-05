using System;
using System.Runtime.Serialization;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport
{
    /// <summary>
    /// This occurs when a value-setting statement has an invalid target - a constant, for example
    /// </summary>
    [Serializable]
    public sealed class IllegalAssignmentException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Illegal assignment";

        [Obsolete("do not use it")] public IllegalAssignmentException() { }
        public IllegalAssignmentException(string additionalInformationIfAny) : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, null) { } // use by generated .cs code! CAUTION: Namespace and Arguments!
        public IllegalAssignmentException(string additionalInformationIfAny, Exception innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 501; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        private IllegalAssignmentException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
