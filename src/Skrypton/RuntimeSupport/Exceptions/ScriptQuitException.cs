using System;

namespace Skrypton.RuntimeSupport
{
    /// <summary>
    /// Ends the script (WScript.Quit) with an exit code - it is not a script error: On Error Resume Next does not trap it, the host that runs the
    /// script catches it (see IsScriptQuit) and ends with the exit code
    /// </summary>
#pragma warning disable CA1032 // a quit always has an exit code
    public sealed class ScriptQuitException : Exception
#pragma warning restore CA1032
    {
        public ScriptQuitException(int exitCode) : base("The script ended with exit code " + exitCode.ToString(System.Globalization.CultureInfo.InvariantCulture))
        {
            ExitCode = exitCode;
        }

        public int ExitCode { get; }

        /// <summary>
        /// The ScriptQuitException in the exception or its inner exceptions (reflection and the script control wrap it), or null
        /// </summary>
        public static ScriptQuitException? Find(Exception? exception)
        {
            for (Exception? e = exception; e != null; e = e.InnerException)
            {
                if (e is ScriptQuitException quit)
                    return quit;
            }
            return null;
        }
    }
}
