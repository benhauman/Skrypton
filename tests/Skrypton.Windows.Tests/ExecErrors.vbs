' WshShell.Exec / Run of a program that can not be started (issue #165): the test runs this script with cscript and the command line tool and
' compares the output - the Windows error as HRESULT (0x80070002), Exec with the source WshShell.Exec and the message of Windows, Run without both
On Error Resume Next
Dim sh : Set sh = CreateObject("WScript.Shell")
Dim e : Set e = sh.Exec("C:\does-not-exist\nothing.exe")
WScript.Echo "exec missing folder: " & Err.Number & " " & Hex(Err.Number) & " [" & Err.Source & "] [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Set e = sh.Exec("nothing-skrypton-xyz.exe")
WScript.Echo "exec missing program: " & Err.Number & " " & Hex(Err.Number) & " [" & Err.Source & "] [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Set e = sh.Exec(sh.ExpandEnvironmentStrings("%windir%") & "\System32\nothing-skrypton-xyz.exe")
WScript.Echo "exec missing program (full path): " & Err.Number & " " & Hex(Err.Number) & " [" & Err.Source & "] [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Dim r : r = sh.Run("C:\does-not-exist\nothing.exe", 0, True)
WScript.Echo "run missing folder: " & Err.Number & " " & Hex(Err.Number) & " [" & Err.Source & "] [" & Err.Description & "]" : Err.Clear
r = sh.Run("nothing-skrypton-xyz.exe", 0, True)
WScript.Echo "run missing program: " & Err.Number & " " & Hex(Err.Number) & " [" & Err.Source & "] [" & Err.Description & "]" : Err.Clear
