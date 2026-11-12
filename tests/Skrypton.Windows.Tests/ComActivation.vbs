' COM activation (issue #103): the test runs this script with cscript and with the command line tool (/com:WScript.Network,System.Collections.ArrayList) and compares the output
Dim net : Set net = CreateObject("WScript.Network")
WScript.Echo TypeName(net) & " " & (Len(net.ComputerName) > 0) & " " & (net.ComputerName = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%"))
Dim list : Set list = CreateObject("System.Collections.ArrayList")
list.Add 3
list.Add "b"
list.Add 1.5
WScript.Echo TypeName(list) & " " & list.Count & " " & list.Item(1) & " " & list(0) & " " & list.Contains("b") & " " & list.IndexOf(1.5, 0)
list.Capacity = 20
WScript.Echo list.Capacity
list.Remove "b"
list.Insert 0, "first"
Dim x, s
For Each x In list
    s = s & x & ";"
Next
WScript.Echo s & " " & list.Count
Class Item
    Public Name
    Public Function Hello()
        Hello = "hello " & Name
    End Function
End Class
Dim it : Set it = New Item
it.Name = "script"
list.Add it
WScript.Echo TypeName(list.Item(list.Count - 1)) & " " & list.Item(list.Count - 1).Hello()
On Error Resume Next
Dim nope : Set nope = CreateObject("Foo.Bar")
WScript.Echo "missing: " & Err.Number & " " & Err.Description
Err.Clear
