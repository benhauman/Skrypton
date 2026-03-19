using System;
using System.Runtime.Serialization;

namespace Skrypton.ScriptControlSupport;

[Serializable]
public sealed class CompilationFailedException : Exception
{
    [Obsolete("do not use it")]
    public CompilationFailedException()
    {
    }
    public CompilationFailedException(string message) : base(message)
    {
    }
    public CompilationFailedException(string message, Exception innerException)
        : base(message, innerException) { }

    private CompilationFailedException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}