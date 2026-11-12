' Parity script: the expected output (ComParity_CallArguments.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' A call statement without brackets whose argument has a "=" - the "=" is a comparison, not an assignment (issue #120)
Sub Two(a, b)
    Out a & "|" & b
End Sub
Out 1 = 1
Out 1 = 2
Two False, 1 = 1
Two 1 = 1, 3
Two 2, 1 < 2
Dim x : x = "a"
Out x = "a"
Out "x" & 1 = "x1"
Call Two(1 = 2, 3)
Dim a : a = 1 = 1
Out a
Dim b(2) : b(1) = 5
Out b(1)
Dim m(1, 1) : m(0, 1) = 3
Out m(0, 1)
Dim d : Set d = CreateObject("Scripting.Dictionary")
d("k") = 1 = 1
Out d("k")
d.Item("j") = 7
Out d.Item("j")
