' Parity script: the expected output (ComParity_VBScriptRegExp.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
Dim re : Set re = CreateObject("VBScript.RegExp")
Out TypeName(re) & "|" & VarType(re)
re.Pattern = "(\w+)@(\w+)\.com"
Out re.Test("mail bob@example.com now") & "|" & re.Test("none")
re.Global = True : re.IgnoreCase = True
Dim ms : Set ms = re.Execute("A@B.com and c@d.COM")
Out TypeName(ms) & "|" & ms.Count
Dim m : Set m = ms(1)
Out TypeName(m) & "|" & m.Value & "|" & m.FirstIndex & "|" & m.Length & "|" & m.SubMatches.Count & "|" & m.SubMatches(0)
Out re.Replace("A@B.com and c@d.COM", "<$2:$1>")
re.Global = False
Out re.Replace("x1 x2 x3", "y")
re.Pattern = "x(\d)" : re.Global = True
Out re.Replace("x1 x2 x3", "[$1]")
Dim s : s = ""
For Each m In re.Execute("x7 x8") : s = s & m.SubMatches(0) : Next
Out s
re.Pattern = "(a)(b)?" : re.Global = False
Set m = re.Execute("ac")(0)
Out m.SubMatches.Count & "|" & IsEmpty(m.SubMatches(1)) & "|" & TypeName(m.SubMatches(1)) & "|" & TypeName(m.SubMatches)
Out VarType(m) & "|" & m
