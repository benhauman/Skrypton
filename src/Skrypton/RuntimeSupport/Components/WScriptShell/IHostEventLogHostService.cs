namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that WshShell.LogEvent writes through (issue #116) - the Windows host service writes to the Application event log with
    /// the source WSH (as Windows Script Host does); a host can redirect it (eg. to its own trace) or switch it off. Without one, the entries go to
    /// System.Diagnostics.Trace.
    /// </summary>
    public interface IHostEventLogHostService
    {
        /// <summary>
        /// type: 0 success, 1 error, 2 warning, 4 information, 8 audit success, 16 audit failure (also the event id, as WSH writes it); target: the
        /// computer whose log is written (null or empty for the local computer)
        /// </summary>
        void LogEvent(int type, string message, string? target);
    }

    internal sealed class DefaultHostEventLogHostService : IHostEventLogHostService
    {
        public static readonly DefaultHostEventLogHostService Instance = new DefaultHostEventLogHostService();

        private DefaultHostEventLogHostService()
        {
        }

        public void LogEvent(int type, string message, string? target)
        {
            System.Diagnostics.Trace.WriteLine(message, "WSH LogEvent " + type.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
