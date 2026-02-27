using System;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.ScriptControlSupport;

namespace Skrypton.RuntimeSupport
{
    public sealed class ErrorDetails : IScriptError
    {
        public static readonly ErrorDetails NoError = new ErrorDetails(0, "", "", "", null);

        public ErrorDetails(int number, string source, string text, string description, Exception originalExceptionIfKnown)
        {
            Number = number;
            Source = source ?? "";
            Text = text ?? "";
            Description = description ?? "";
            OriginalExceptionIfKnown = originalExceptionIfKnown;
        }

        /// <summary>
        /// This will be zero if there is no current error
        /// </summary>
        [IsDefault]
        public int Number { get; private set; }

        /// <summary>
        /// This will be a blank string if there is no current error
        /// </summary>
        public string Source { get; private set; }

        /// <summary>
        /// This will be a blank string if there is no current error
        /// </summary>
        public string Description { get; private set; }
        public string Text { get; private set; }

        /// <summary>
        /// This will be non-null Number is non-zero and if the exception was caught and translated into an ErrorDetails instance, but it may be null even
        /// if Number is non-zero (if the error was created with a RAISEERROR call)
        /// </summary>
        public Exception OriginalExceptionIfKnown { get; private set; }

        public string HelpFile { get; }
        public int HelpContext { get; }
        public int Line { get; }
        public int Column { get; }
        public void Clear()
        {
            //!?!?!
        }
    }
}
