' Parity script: the expected output (ComParity_OutlookApplication.expected.txt) was captured with Windows cscript.exe and Microsoft Outlook.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' Nothing is saved, sent or displayed (the items stay drafts in memory); localized names (folders) are not compared.
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim x
Dim o : Set o = CreateObject("Outlook.Application")
Out TypeName(o) & "|" & VarType(o) & "|" & o & "|" & o.Name

Dim m : Set m = o.CreateItem(0)
Out TypeName(m) & "|" & VarType(m) & "|" & m.Class & "|[" & m.Subject & "]|[" & m.To & "]|" & m.Importance & "|" & m.BodyFormat & "|" & m.Sent & "|" & m.Saved
m.Subject = "Betreff" : m.Body = "Text" : m.CC = "c@example.invalid"
m.To = "a@example.invalid; b@example.invalid"
Out m & "|" & m.Subject & "|" & m.To & "|" & m.CC & "|" & m.Recipients.Count & "|" & TypeName(m.Recipients)
Dim r : Set r = m.Recipients.Add("d@example.invalid")
Out TypeName(r) & "|" & r & "|" & r.Type & "|" & m.Recipients.Count & "|" & m.To & "|" & m.Recipients(1).Name
m.HTMLBody = "<b>fett</b>"
Out m.BodyFormat & "|" & (InStr(m.HTMLBody, "<b>fett</b>") > 0)
Dim file : file = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton_parity_" & fso.GetTempName())
fso.CreateTextFile(file).Close
Dim at : Set at = m.Attachments.Add(file)
Out TypeName(m.Attachments) & "|" & m.Attachments.Count & "|" & TypeName(at) & "|" & (at.FileName = fso.GetFileName(file))

Dim c : Set c = o.CreateItem(2)
Out TypeName(c) & "|" & VarType(c) & "|" & c.Class & "|[" & c.FirstName & "]|[" & c.FullName & "]|" & TypeName(c.Birthday) & "|" & Year(c.Birthday)
c.FirstName = "Max" : c.LastName = "Muster" : c.CompanyName = "helpLine" : c.Email1Address = "max@example.invalid"
Out c & "|" & c.FullName & "|" & c.FileAs & "|" & c.CompanyName & "|" & c.Email1Address

Dim a : Set a = o.CreateItem(1)
Out TypeName(a) & "|" & VarType(a) & "|" & a.Class & "|" & a.Duration & "|" & a.ReminderSet & "|" & a.ReminderMinutesBeforeStart & "|" & a.AllDayEvent & "|" & a.BusyStatus & "|" & TypeName(a.Start) & "|" & DateDiff("n", a.Start, a.End)
a.Start = DateSerial(2026, 11, 12) + TimeSerial(10, 0, 0)
a.Duration = 90
Out Hour(a.End) & ":" & Minute(a.End) & "|" & TypeName(a.End) & "|" & Year(a.Start) & "-" & Month(a.Start) & "-" & Day(a.Start) & " " & Hour(a.Start)
a.End = DateSerial(2026, 11, 12) + TimeSerial(12, 0, 0)
a.Subject = "Termin" : a.Location = "Raum 1"
Out a.Duration & "|" & a.Subject & "|" & a.Location

Dim ns : Set ns = o.GetNamespace("MAPI")
Out TypeName(ns) & "|" & VarType(ns) & "|" & ns & "|" & ns.Type & "|" & TypeName(o.Session)
Dim f : Set f = ns.GetDefaultFolder(10)
Out TypeName(f) & "|" & VarType(f) & "|" & f.DefaultItemType & "|" & TypeName(f.Items) & "|" & VarType(f.Items)
Out TypeName(f.Items.Find("[LastName] = 'NoSuchSkryptonName' And [FirstName] = 'Nobody'"))
Set x = ns.GetDefaultFolder(9).Items.Add()
Out TypeName(x) & "|" & ns.GetDefaultFolder(9).DefaultItemType
Set x = ns.CreateRecipient("nobody-skrypton")
Out TypeName(x) & "|" & x.Name & "|" & x.Resolved

On Error Resume Next
Set x = o.CreateItem(99) : Out "item 99:" & Err.Number : Err.Clear
Set x = o.GetNamespace("XYZ") : Out "namespace XYZ:" & Err.Number : Err.Clear
Set x = m.Attachments.Add(file & ".missing") : Out "missing attachment:" & (Err.Number <> 0) : Err.Clear
On Error GoTo 0
fso.DeleteFile file
