' Parity script: the expected output (ComParity_WScriptShellScriptletWmiAdodbTypeNames.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
Dim sh : Set sh = CreateObject("WScript.Shell")
Out TypeName(sh) & "|" & VarType(sh)
Out sh.ExpandEnvironmentStrings("%SKRYPTON_PARITY_VAR%|%NO_SUCH_VAR_X%")
Dim tl : Set tl = CreateObject("Scriptlet.TypeLib")
Dim g : g = tl.Guid
Out TypeName(tl) & "|" & Len(g) & "|" & Left(g, 1) & "|" & Mid(g, 38, 1) & "|" & (Mid(g, 10, 1) = "-")
Dim xh : Set xh = CreateObject("Msxml2.ServerXMLHTTP.6.0")
Out TypeName(xh) & "|" & xh.readyState
Dim app : Set app = CreateObject("Shell.Application")
Out TypeName(app)
Dim wmi : Set wmi = GetObject("winmgmts:")
Out TypeName(wmi)
Dim ps : Set ps = wmi.InstancesOf("Win32_Process")
Out TypeName(ps) & "|" & (ps.Count > 0)
Dim p : For Each p In ps : Out TypeName(p) & "|" & (Len(p.Name) > 0) & "|" & VarType(p.ProcessId) : Exit For : Next
Out TypeName(CreateObject("WbemScripting.SWbemLocator"))
Out TypeName(CreateObject("ADODB.Connection")) & "|" & TypeName(CreateObject("ADODB.Recordset")) & "|" & TypeName(CreateObject("ADODB.Command")) & "|" & TypeName(CreateObject("ADODB.Stream"))
Out VarType(CreateObject("ADODB.Recordset")) & "|" & VarType(CreateObject("ADODB.Command")) & "|" & VarType(CreateObject("ADODB.Stream")) & "|" & VarType(CreateObject("Scriptlet.TypeLib")) & "|" & VarType(CreateObject("Msxml2.DOMDocument"))
