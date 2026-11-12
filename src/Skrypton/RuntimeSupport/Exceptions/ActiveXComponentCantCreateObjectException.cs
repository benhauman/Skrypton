using System;
using System.Runtime.Serialization;

namespace Skrypton.RuntimeSupport.Exceptions
{
    /// <summary>
    /// Error 429: CreateObject of a ProgID that nothing creates (no factory, no built-in, no COM class the host allows) - VBScript names the ProgID
    /// when it reports the error to the host ("ActiveX component can't create object: 'Foo.Bar'", verified with cscript); Err.Description in the
    /// script is "ActiveX component can't create object" (issue #103)
    /// </summary>
    [Serializable]
    public sealed class ActiveXComponentCantCreateObjectException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "ActiveX component can't create object";

        [Obsolete("do not use it")] private ActiveXComponentCantCreateObjectException() : this(null, innerException: null) { }
        public ActiveXComponentCantCreateObjectException(string message) : this(message, innerException: null) { }
        public ActiveXComponentCantCreateObjectException(string? additionalInformationIfAny, Exception? innerException)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public static ActiveXComponentCantCreateObjectException ForProgId(string progId, string? additionalInformationIfAny)
        {
            return new ActiveXComponentCantCreateObjectException(additionalInformationIfAny, innerException: null) { HostTarget = progId };
        }

        public override int ErrorNumber { get { return 429; } }

        private ActiveXComponentCantCreateObjectException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
