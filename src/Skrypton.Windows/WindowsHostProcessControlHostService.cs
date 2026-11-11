using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// Starts and controls processes of Windows (WshShell.Run / Exec / AppActivate, Shell.ShellExecute, Win32_Process)
    /// </summary>
    public sealed class WindowsHostProcessControlHostService : IHostProcessControlHostService
    {
        private static readonly char[] Whitespace = { ' ', '\t' };

        public int ProcessStart(string command, byte windowMode, bool waitOnReturn)
        {
            (string file, string arguments) = SplitCommand(Environment.ExpandEnvironmentVariables(command ?? ""));
            var startInfo = new ProcessStartInfo(file, arguments) { UseShellExecute = true, WindowStyle = ToWindowStyle(windowMode) };
            using (Process? process = Process.Start(startInfo))
            {
                if (!waitOnReturn || (process == null))
                    return 0;
                process.WaitForExit();
                return process.ExitCode;
            }
        }

        public bool ProcessActivate(int processId)
        {
            Process process;
            try
            {
                process = Process.GetProcessById(processId);
            }
            catch (ArgumentException)
            {
                return false;
            }
            using (process)
            {
                IntPtr window = process.MainWindowHandle;
                return (window != IntPtr.Zero) && NativeMethods.SetForegroundWindow(window);
            }
        }

        public void ProcessesCollect(Func<int, string, bool> collector)
        {
            if (collector == null)
                throw new ArgumentNullException(nameof(collector));
            foreach (Process process in Process.GetProcesses().OrderBy(p => p.Id))
            {
                using (process)
                {
                    // Win32_Process.Name is the image name ("notepad.exe")
                    string name = (process.Id == 0) ? "System Idle Process" : (process.Id == 4) ? "System" : process.ProcessName + ".exe";
                    if (!collector(process.Id, name))
                        return;
                }
            }
        }

        public HostProcessExecResult ProcessExec(string command)
        {
            (string file, string arguments) = SplitCommand(Environment.ExpandEnvironmentVariables(command ?? ""));
            var startInfo = new ProcessStartInfo(file, arguments)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            using (Process process = Process.Start(startInfo) ?? throw new InvalidOperationException($"WshShell.Exec: '{command}' could not be started"))
            {
                // read both streams while the process runs (a full pipe would block it)
                var standardError = process.StandardError.ReadToEndAsync();
                string standardOutput = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                return new HostProcessExecResult(process.Id, process.ExitCode, standardOutput, standardError.Result);
            }
        }

        public void ShellExecute(string file, string arguments, string directory, string operation, int show)
        {
            var startInfo = new ProcessStartInfo(file ?? "", arguments ?? "")
            {
                UseShellExecute = true,
                Verb = operation ?? "",
                WindowStyle = ToWindowStyle(show),
            };
            if (!string.IsNullOrEmpty(directory))
                startInfo.WorkingDirectory = directory;
            Process.Start(startInfo)?.Dispose();
        }

        /// <summary>
        /// "notepad.exe a.txt" / "\"C:\Program Files\x.exe\" /a" -> file, arguments
        /// </summary>
        internal static (string File, string Arguments) SplitCommand(string command)
        {
            string text = command.TrimStart();
            if (text.StartsWith("\"", StringComparison.Ordinal))
            {
                int end = text.IndexOf('"', 1);
                if (end > 0)
                    return (text.Substring(1, end - 1), text.Substring(end + 1).TrimStart());
            }
            int space = text.IndexOfAny(Whitespace);
            return (space < 0) ? (text, "") : (text.Substring(0, space), text.Substring(space + 1).TrimStart());
        }

        /// <summary>
        /// The window style of WshShell.Run / ShowWindow: 0 hidden, 2 / 6 / 7 minimized, 3 maximized, otherwise normal
        /// </summary>
        internal static ProcessWindowStyle ToWindowStyle(int windowMode)
        {
            switch (windowMode)
            {
                case 0:
                    return ProcessWindowStyle.Hidden;
                case 2:
                case 6:
                case 7:
                    return ProcessWindowStyle.Minimized;
                case 3:
                    return ProcessWindowStyle.Maximized;
                default:
                    return ProcessWindowStyle.Normal;
            }
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetForegroundWindow(IntPtr hWnd);
        }
    }
}
