' Parity script: the expected output (ComParity_Scopes.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' A Dim / Const in a Sub / Function may have the name of a global declared further down (issue #130)
Sub S()
    Dim x : x = 1
    Out "local " & x
End Sub
Dim x : x = 2
S
Out "global " & x
Function F()
    Dim Y : Y = "fy"
    F = Y
End Function
Dim y : y = "gy"
Out F() & " " & y
Sub C()
    Const z = "local const"
    Out z
End Sub
C
Const z = "global const"
Out z
