' WshShell.Environment (issue #107): the test runs this script with cscript and the command line tool and compares the output (it depends on the environment of the machine)
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim e : Set e = oShell.Environment("PROCESS")
WScript.Echo TypeName(e) & " " & (e("ProgramFiles") = oShell.ExpandEnvironmentStrings("%ProgramFiles%")) & " [" & e("NO_SUCH_VARIABLE") & "]"
e("SKRYPTON_PROBE") = "1"
WScript.Echo oShell.ExpandEnvironmentStrings("%SKRYPTON_PROBE%") & " " & (oShell.Environment("PROCESS").Count > 0) & " " & e.Item("SKRYPTON_PROBE") & " " & e.Item("skrypton_probe") & " " & (e.Count = e.length)
Dim v, n, found, eq : n = 0
For Each v In e
    n = n + 1
    If v = "SKRYPTON_PROBE=1" Then found = True
    If Left(v, 1) = "=" Then eq = eq + 1
Next
WScript.Echo "n=count " & (n = e.Count) & " found " & found & " eq " & eq
e.Remove "SKRYPTON_PROBE"
WScript.Echo "[" & e("SKRYPTON_PROBE") & "] " & oShell.ExpandEnvironmentStrings("%SKRYPTON_PROBE%")
On Error Resume Next
e.Remove "SKRYPTON_NOT_THERE"
WScript.Echo "remove missing: " & Err.Number & " " & Err.Description : Err.Clear
Dim s : Set s = oShell.Environment
WScript.Echo "noarg " & TypeName(s) & " " & (s("OS") = oShell.Environment("SYSTEM")("OS")) & " [" & s("OS") & "]"
WScript.Echo "SYSTEM path " & (Len(oShell.Environment("SYSTEM")("Path")) > 0) & " " & (InStr(oShell.Environment("SYSTEM")("Path"), "%") > 0)
WScript.Echo "USER temp [" & oShell.Environment("USER")("TEMP") & "]"
WScript.Echo "VOLATILE " & (Len(oShell.Environment("VOLATILE")("USERNAME")) > 0)
Dim bad : Set bad = oShell.Environment("BOGUS")
WScript.Echo "bogus: " & Err.Number & " " & Err.Description : Err.Clear
WScript.Echo "process case " & (oShell.Environment("process")("OS") <> "")
e("") = "x"
WScript.Echo "empty name: " & Err.Number & " " & Err.Description : Err.Clear
