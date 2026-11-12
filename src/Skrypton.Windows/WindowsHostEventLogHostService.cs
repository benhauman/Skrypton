using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// WshShell.LogEvent as Windows Script Host writes it (issue #116): an entry of the Application event log with the source WSH, the type as the
    /// entry type and as the event id with the severity bits of WSH (issue #139 - ReportEvent: EventLog.WriteEntry does not accept the success type 0)
    /// </summary>
    public sealed class WindowsHostEventLogHostService : IHostEventLogHostService
    {
        private const string Source = "WSH";

        public void LogEvent(int type, string message, string? target)
        {
            string? server = (string.IsNullOrEmpty(target) || (target == ".")) ? null : target;
            IntPtr eventSource = RegisterEventSource(server, Source);
            if (eventSource == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(), $"WshShell.LogEvent: the event log of '{target ?? "."}' can not be opened");
            try
            {
                if (!ReportEvent(eventSource, (ushort)type, 0, InstanceId(type), IntPtr.Zero, 1, 0, new[] { message ?? "" }, IntPtr.Zero))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "WshShell.LogEvent: the entry can not be written");
            }
            finally
            {
                _ = DeregisterEventSource(eventSource);
            }
        }

        /// <summary>
        /// The event id with the severity bits of the message of the WSH message file (wshom.ocx), as Windows Script Host writes it - without
        /// them the event log can not find the message (issue #139): error and audit failure 0xC0000000, warning 0x80000000, information 0x40000000
        /// </summary>
        private static uint InstanceId(int type)
        {
            switch (type)
            {
                case 1:
                case 16:
                    return 0xC0000000u | (uint)type;
                case 2:
                    return 0x80000000u | (uint)type;
                case 4:
                    return 0x40000000u | (uint)type;
                default:
                    return (uint)type;
            }
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "RegisterEventSourceW")]
        private static extern IntPtr RegisterEventSource(string? lpUNCServerName, string lpSourceName);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "ReportEventW")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ReportEvent(IntPtr hEventLog, ushort wType, ushort wCategory, uint dwEventID, IntPtr lpUserSid, ushort wNumStrings, uint dwDataSize, string[] lpStrings, IntPtr lpRawData);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeregisterEventSource(IntPtr hEventLog);
    }
}
