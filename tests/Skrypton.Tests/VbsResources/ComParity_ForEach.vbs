' Parity script: the expected output (ComParity_ForEach.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' At the end of a For Each loop the loop variable is Empty; with Exit For it keeps its value (issue #112)
Dim v, d
For Each v In Array(1, 2)
Next
Out "array " & TypeName(v)
Set d = CreateObject("Scripting.Dictionary") : d.Add "a", 1
For Each v In d
Next
Out "dict " & TypeName(v)
For Each v In Array(1, 2)
If v = 1 Then Exit For
Next
Out "exit " & TypeName(v) & " " & v
For Each v In Array()
Next
Out "empty " & TypeName(v)
Dim i
For i = 1 To 2
Next
Out "for " & i
' with error trapping, a target that cannot be enumerated enters the loop once and the variable keeps its value
On Error Resume Next
Dim n : v = 5
For Each v In "abc"
n = n + 1
Next
Out "string " & TypeName(v) & " " & v & " n=" & n & " err=" & Err.Number
Err.Clear
v = 7 : n = 0
For Each v In Nothing
n = n + 1
Next
Out "nothing " & TypeName(v) & " " & v & " n=" & n
Err.Clear
For Each v In Array(1)
Next
Out "ok " & TypeName(v)
Sub S(x)
For Each x In Array(1, 2)
Next
End Sub
Dim y : y = 9 : S y
Out "byref " & TypeName(y)
