' Demo: run commands and look at processes - WScript.Shell (Run, Exec, environment) and WMI (winmgmts:).
'   Skrypton.Windows.Cli Shell-Processes.vbs
Option Explicit
Dim shell, execution, exitCode, wmi, processes, process, count
Set shell = CreateObject("WScript.Shell")
WScript.Echo "OS: " & shell.ExpandEnvironmentStrings("%OS%")

Set execution = shell.Exec("cmd /c echo Hello from cmd")
Do While execution.Status = 0
    WScript.Sleep 50
Loop
WScript.Echo "Exec output: " & Trim(Replace(execution.StdOut.ReadAll(), vbCrLf, ""))
WScript.Echo "Exec exit code: " & execution.ExitCode

exitCode = shell.Run("cmd /c exit 3", 0, True) ' hidden, wait for the exit code
WScript.Echo "Run exit code: " & exitCode

Set wmi = GetObject("winmgmts:")
Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'System'")
count = 0
For Each process In processes
    count = count + 1
Next
WScript.Echo "System process found: " & (count > 0)
