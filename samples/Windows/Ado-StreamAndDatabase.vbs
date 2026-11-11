' Demo: ADO - write and read a UTF-8 text file with ADODB.Stream; with a connection string also query a database (SQL Server).
'   Skrypton.Windows.Cli Ado-StreamAndDatabase.vbs
'   Skrypton.Windows.Cli Ado-StreamAndDatabase.vbs "Provider=MSOLEDBSQL;Data Source=.;Initial Catalog=master;Integrated Security=SSPI"
Option Explicit
Const adTypeText = 2
Const adSaveCreateOverWrite = 2
Dim fso, file, stream, text, connection, recordset, field
Set fso = CreateObject("Scripting.FileSystemObject")
file = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-demo-ado.txt")

Set stream = CreateObject("ADODB.Stream")
stream.Type = adTypeText
stream.Charset = "utf-8"
stream.Open
stream.WriteText "Gr" & ChrW(252) & ChrW(223) & "e aus Skrypton" & vbCrLf & "Zeile 2"
stream.SaveToFile file, adSaveCreateOverWrite
stream.Close
WScript.Echo "Written: " & fso.GetFile(file).Size & " bytes (with the byte order mark)"

Set stream = CreateObject("ADODB.Stream")
stream.Type = adTypeText
stream.Charset = "utf-8"
stream.Open
stream.LoadFromFile file
text = stream.ReadText
stream.Close
WScript.Echo "Read: " & Len(text) & " characters, line 1 has " & Len(Split(text, vbCrLf)(0)) & ", line 2 is " & Split(text, vbCrLf)(1)
fso.DeleteFile file

If WScript.Arguments.Count > 0 Then
    Set connection = CreateObject("ADODB.Connection")
    connection.Open WScript.Arguments(0)
    Set recordset = connection.Execute("SELECT DB_NAME() AS DatabaseName, SUSER_SNAME() AS LoginName, GETDATE() AS ServerTime")
    For Each field In recordset.Fields
        WScript.Echo "  " & field.Name & " = " & field.Value
    Next
    recordset.Close
    connection.Close
Else
    WScript.Echo "No connection string: the database part is skipped"
End If
