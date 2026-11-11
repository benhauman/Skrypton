using System;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in WScript.Shell (Run, Exec, AppActivate) and process queries go through. Starting processes from a script
    /// is a decision for the host, so there is no default implementation: without a registration these members raise an error.
    /// </summary>
    public interface IHostProcessControlHostService
    {
        /// <summary>
        /// Start the command (WshShell.Run). windowMode is the WshShell.Run window style (0 hidden, 1 normal, 2 minimized, 3 maximized, ...). If
        /// waitOnReturn is true then this waits for the process to exit and returns its exit code, otherwise it returns 0 straight away.
        /// </summary>
        int ProcessStart(string command, byte windowMode, bool waitOnReturn);

        /// <summary>
        /// Activate the window of the process (WshShell.AppActivate) - false if there is no such process
        /// </summary>
        bool ProcessActivate(int processId);

        /// <summary>
        /// Report the running processes (process id, name) to the collector until it returns false
        /// </summary>
        void ProcessesCollect(Func<int, string, bool> collector);

        /// <summary>
        /// Run the command line application, wait for it to exit and return its exit code and output (WshShell.Exec)
        /// </summary>
        HostProcessExecResult ProcessExec(string command);

        /// <summary>
        /// Open a file, document or URL with its associated application (Shell.Application.ShellExecute) - operation is the verb ("open", "print",
        /// "edit", .. or empty for the default) and show is the ShowWindow style (1 = normal)
        /// </summary>
        void ShellExecute(string file, string arguments, string directory, string operation, int show);
    }

    public sealed class HostProcessExecResult
    {
        public HostProcessExecResult(int processId, int exitCode, string stdOut, string stdErr)
        {
            ProcessId = processId;
            ExitCode = exitCode;
            StdOut = stdOut ?? "";
            StdErr = stdErr ?? "";
        }

        public int ProcessId { get; }
        public int ExitCode { get; }
        public string StdOut { get; }
        public string StdErr { get; }
    }

    /// <summary>
    /// The host service that the built-in WScript.Shell resolves environment variables through (ExpandEnvironmentStrings). If the host has not
    /// registered one, the process environment is used (Environment.ExpandEnvironmentVariables).
    /// </summary>
    public interface IHostEnvironmentHostService
    {
        string ExpandEnvironmentStrings(string source);
    }

    internal sealed class DefaultHostEnvironmentHostService : IHostEnvironmentHostService
    {
        public static readonly DefaultHostEnvironmentHostService Instance = new DefaultHostEnvironmentHostService();

        private DefaultHostEnvironmentHostService()
        {
        }

        public string ExpandEnvironmentStrings(string source) => Environment.ExpandEnvironmentVariables(source ?? "");
    }
}
