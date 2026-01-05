using System;

namespace Skrypton.RuntimeSupport.Implementations
{
    public interface IRuntimeLogger
    {
        void LogException(Exception exception);
    }
}