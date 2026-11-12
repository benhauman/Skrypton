' Parity script: the expected output (ComParity_FunctionNameMemberAccess.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' Inside a Function (or Property Get) its name is the return value: Name.Member and Name.Member args use the object assigned to it, they do not
' call the function again (issue #161 - that recursed until the stack overflowed). Name(args) of a parameterless function is error 450.
Function hlITIL()
    Set hlITIL = CreateObject("Scripting.Dictionary")
    hlITIL.Add "key", "value"
    Out "inside: " & hlITIL.Count & " " & hlITIL.Item("key")
    Dim x : x = hlITIL.Count
    hlITIL.Item("k2") = "v2"
    Out "x=" & x & " " & hlITIL.Count & " " & hlITIL.Exists("k2")
End Function
Out "outside: " & hlITIL.Count & " " & hlITIL().Exists("key")

Function Plain()
    Plain = 5
    Out "plain " & (Plain + 1)
End Function
Out Plain()

Class CProp
    Property Get Obj()
        Set Obj = CreateObject("Scripting.Dictionary")
        Obj.Add "p", 1
        Out "property " & Obj.Count
    End Property
    Function G()
        Set G = CreateObject("Scripting.Dictionary")
        G.Add "k", "v"
    End Function
End Class
Dim cp : Set cp = New CProp
Out cp.Obj.Count

Function Rec(n)
    If n = 0 Then Rec = 0 : Exit Function
    Rec = n + Rec(n - 1)
End Function
Out Rec(4)

Function D()
    Set D = CreateObject("Scripting.Dictionary")
    D.Add "key", "value"
    On Error Resume Next
    Dim y : y = D("key")
    Out "y = D(""key""): " & Err.Number & " " & Err.Description : Err.Clear
End Function
Dim r : Set r = D()
Out TypeName(r) & " " & r.Count

On Error Resume Next
Dim z : z = r.Count
z = D("key")
Out "outside D(""key""): " & Err.Number & " " & Err.Description : Err.Clear
z = cp.G("k")
Out "cp.G(""k""): " & Err.Number & " " & Err.Description : Err.Clear
