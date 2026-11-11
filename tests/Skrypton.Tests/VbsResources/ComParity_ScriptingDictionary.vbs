' Parity script: the expected output (ComParity_ScriptingDictionary.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
Dim d : Set d = CreateObject("Scripting.Dictionary")
Out TypeName(d) & "|" & VarType(d)
d.Add "a", 1 : d.Add "b", 2
Out d.Count & "|" & d.Exists("a") & "|" & d.Exists("A")
Out d("missing") & "|" & d.Count & "|" & IsEmpty(d("missing"))
Out Join(d.Keys, ",") & "|" & Join(d.Items, ",")
d.Key("b") = "c"
Out Join(d.Keys, ",")
On Error Resume Next
d.Add "a", 9 : Out "dup:" & Err.Number : Err.Clear
d.Remove "zzz" : Out "remove:" & Err.Number : Err.Clear
d.Key("zzz") = "y" : Out "key:" & Err.Number : Err.Clear
d.CompareMode = 1 : Out "cmp:" & Err.Number : Err.Clear
On Error GoTo 0
Dim e : Set e = CreateObject("Scripting.Dictionary")
e.CompareMode = 1 : e.Add "X", 1 : Out e.Exists("x")
e.Add 1, "n" : e.Add "1", "s" : Out e.Count & "|" & e(1) & "|" & e("1")
e.RemoveAll : Out e.Count
