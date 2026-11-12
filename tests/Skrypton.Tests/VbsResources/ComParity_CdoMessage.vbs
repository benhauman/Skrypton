' Parity script: the expected output (ComParity_CdoMessage.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' CDO.Configuration / CDO.Message (CDOSYS, issue #149) without a mail server: the fields, the bodies and the errors of Send
Const cdo = "http://schemas.microsoft.com/cdo/configuration/"
Dim c : Set c = CreateObject("CDO.Configuration")
Out TypeName(c) & " " & TypeName(c.Fields) & " " & TypeName(c.Fields.Item(cdo & "sendusing"))
Out "defaults [" & c.Fields(cdo & "languagecode") & "] [" & c.Fields(cdo & "postusing") & "] [" & c.Fields(cdo & "sendusing") & "] [" & c.Fields(cdo & "usemessageresponsetext") & "] [" & c.Fields(cdo & "smtpserver") & "]"
c.Fields(cdo & "smtpserver") = "mail.example.com"
c.Fields.Item(cdo & "smtpserverport").Value = 2525
Out "set [" & c.Fields(cdo & "smtpserver") & "] [" & c.Fields.Item(cdo & "smtpserverport").Value & "] [" & c.Fields("HTTP://SCHEMAS.MICROSOFT.COM/cdo/configuration/SMTPSERVER") & "] " & c.Fields.Item(cdo & "smtpserver").Name
c.Fields.Update
Dim m : Set m = CreateObject("CDO.Message")
Out TypeName(m) & " " & TypeName(m.Configuration) & " [" & m.Configuration.Fields(cdo & "sendusing") & "] [" & m.From & "|" & m.To & "|" & m.Subject & "|" & m.TextBody & "|" & m.HTMLBody & "]"
m.HTMLBody = "<p>The incident <b>4711</b> is solved.</p><p>Line&nbsp;2 &amp; more<br>next</p>"
Out "text [" & Replace(Replace(m.TextBody, vbCr, "\r"), vbLf, "\n") & "]"
m.TextBody = "plain"
Out "text [" & m.TextBody & "] html [" & m.HTMLBody & "]"
On Error Resume Next
m.From = "a@example.com" : m.To = "b@example.com"
m.Send
Out "no sendusing: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "] " & Err.Source : Err.Clear
m.Configuration.Fields(cdo & "sendusing") = 2
m.Configuration.Fields.Update
m.Send
Out "no server: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Set m.Configuration = c
c.Fields(cdo & "sendusing") = 2
c.Fields.Update
m.From = ""
m.Send
Out "no from: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
m.From = "a@example.com" : m.To = ""
m.Send
Out "no recipient: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Dim m3 : Set m3 = CreateObject("CDO.Message")
m3.Configuration.Fields(cdo & "sendusing") = 2
m3.Configuration.Fields(cdo & "smtpserver") = "127.0.0.1"
m3.Configuration.Fields(cdo & "smtpserverport") = 1
m3.Configuration.Fields(cdo & "smtpconnectiontimeout") = 5
m3.Configuration.Fields.Update
m3.From = "a@example.com" : m3.To = "b@example.com" : m3.TextBody = "t"
m3.Send
Out "no listener: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "] " & Err.Source
