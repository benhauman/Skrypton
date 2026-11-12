' Parity script: the expected output (ComParity_Classes.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' VBScript classes: fields, Class_Initialize / Class_Terminate, Property Get / Let, a default function, a method; an instance in a Dictionary
Class Counter
    Private m_count
    Public Name
    Private Sub Class_Initialize()
        m_count = 10
    End Sub
    Private Sub Class_Terminate()
    End Sub
    Public Property Get Count()
        Count = m_count
    End Property
    Public Property Let Count(value)
        m_count = value
    End Property
    Public Default Function Describe()
        Describe = Name & "=" & m_count
    End Function
    Public Sub Add(n)
        m_count = m_count + n
    End Sub
End Class
Dim c : Set c = New Counter
c.Name = "c"
c.Add 5
Out TypeName(c) & " " & c.Count & " " & c.Describe() & " " & c
c.Count = 1
Out c.Count & " " & VarType(c) & " " & IsObject(c)
Dim d : Set d = CreateObject("Scripting.Dictionary")
d.Add "k", c
Out d("k").Count
