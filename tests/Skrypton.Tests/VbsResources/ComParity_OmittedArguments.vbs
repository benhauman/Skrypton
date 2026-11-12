' Parity script: the expected output (ComParity_OmittedArguments.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' Omitted arguments (issue #154): calls with and without brackets, the first argument left out, TypeName / VarType of the omitted argument,
' omitted optional arguments of a built-in
Class C
    Sub M(a, b, c) : Out TypeName(a) & "|" & TypeName(b) & "|" & c : End Sub
End Class
Dim x : Set x = New C
Sub S(a, b, c) : Out TypeName(a) & "|" & TypeName(b) & "|" & c : End Sub
S 1, , 3
S , , 3
x.M , , 3
x.M 1, , 3
Call x.M(, , 3)
Call S(1, , 3)
Sub V(a, b, c) : Out VarType(b) & " " & IsEmpty(b) & " " & IsNull(b) & " " & IsObject(b) & " " & (VarType(b) = vbError) : End Sub
V 1, , 3
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim p : p = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-omitted-arguments.txt")
If fso.FileExists(p) Then fso.DeleteFile p
Dim f : Set f = fso.OpenTextFile(p, 2, True)
f.Write "abc"
f.Close
Set f = fso.OpenTextFile(p, , False)
Out "OpenTextFile(p, , False): " & f.ReadAll
f.Close
fso.DeleteFile p
