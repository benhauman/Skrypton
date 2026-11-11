' Parity script: the expected output (ComParity_WordApplication.expected.txt) was captured with Windows cscript.exe and Microsoft Word.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' SKRYPTON_PARITY_DOCX is a copy of ComParity_WordApplication.template.docx (created with Word: paragraphs with a bookmark on "XXX" and a
' collapsed bookmark, a 2x3 table with a bookmark on a cell).
Function Show(s) : Show = Replace(Replace(Replace(Replace(Replace(s, vbCr, "\r"), vbLf, "\n"), Chr(7), "\a"), vbTab, "\t"), Chr(11), "\v") : End Function
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim x
Dim sh : Set sh = CreateObject("WScript.Shell")
Dim template : template = sh.ExpandEnvironmentStrings("%SKRYPTON_PARITY_DOCX%")
Dim target : target = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton_parity_" & fso.GetTempName() & ".docx")

Dim w : Set w = CreateObject("Word.Application")
Out TypeName(w) & "|" & VarType(w) & "|" & w & "|" & w.Name & "|" & w.Visible
w.Caption = "Angebot"
Out w.Caption & "|" & TypeName(w.Documents) & "|" & w.Documents.Count

' a template filled in, as scripts do
Dim d : Set d = w.Documents.Open(template)
Out TypeName(d) & "|" & VarType(d) & "|" & (d.Name = fso.GetFileName(template)) & "|" & (d = d.Name) & "|" & w.Documents.Count & "|" & (w.ActiveDocument.Name = d.Name)
Out Show(d.Content.Text)
Out TypeName(d.Bookmarks) & "|" & d.Bookmarks.Count & "|" & d.Bookmarks(1).Name & "|" & d.Bookmarks(2).Name & "|" & d.Bookmarks(3).Name & "|" & d.Bookmarks.Exists("FIRMENNAME") & "|" & d.Bookmarks.Exists("nope")
Dim bm : Set bm = d.Bookmarks("firmenname")
Out TypeName(bm) & "|" & VarType(bm) & "|" & bm & "|" & TypeName(bm.Range) & "|" & VarType(bm.Range) & "|[" & Show(bm.Range.Text) & "]|[" & Show(d.Bookmarks("Vorname").Range.Text) & "]|[" & Show(d.Bookmarks("zelle").Range) & "]"
Dim t : Set t = d.Tables(1)
Out TypeName(d.Tables) & "|" & d.Tables.Count & "|" & TypeName(t) & "|" & VarType(t) & "|" & TypeName(t.Rows) & "|" & t.Rows.Count & "|" & TypeName(t.Rows(1)) & "|" & TypeName(t.Rows(1).Cells) & "|" & t.Rows(1).Cells.Count
Out TypeName(t.Cell(1, 2)) & "|" & VarType(t.Cell(1, 2)) & "|" & TypeName(t.Cell(1, 2).Range) & "|" & Show(t.Cell(1, 2).Range.Text) & "|" & Show(t.Rows(2).Cells(1).Range) & "|" & t.Cell(2, 3).RowIndex & "," & t.Cell(2, 3).ColumnIndex
t.Rows(2).Cells(1).Range = "DL-" & 1
t.Rows(2).Cells(2).Range.Text = "Dienstleistung" & vbCrLf & "(Zeitstunden)"
t.Cell(2, 3).Range.Text = 5055
Out Show(t.Rows(2).Cells(1).Range.Text) & "|" & Show(t.Cell(2, 2).Range.Text) & "|" & Show(t.Cell(2, 3).Range.Text) & "|" & Show(d.Bookmarks("zelle").Range.Text)
t.Rows.Add
Out t.Rows.Count & "|" & Show(t.Cell(3, 1).Range.Text) & "|" & t.Rows(3).Index
Dim r : Set r = d.Bookmarks("firmenname").Range
r.Text = "helpLine"
Out Show(d.Content.Text) & "|" & d.Bookmarks.Exists("firmenname") & "|" & d.Bookmarks.Count
d.Bookmarks("Vorname").Range.Text = "Herr Max"
Out Show(d.Content.Text) & "|" & d.Bookmarks.Exists("Vorname")
d.Bookmarks("zelle").Range.Text = "Zelle"
Out Show(t.Cell(2, 2).Range.Text) & "|" & d.Bookmarks.Exists("zelle") & "|" & d.Saved
On Error Resume Next
Set x = d.Bookmarks("missing") : Out "missing bookmark:" & Err.Number : Err.Clear
Set x = d.Tables(2) : Out "missing table:" & Err.Number : Err.Clear
Set x = t.Rows(0) : Out "row 0:" & Err.Number : Err.Clear
Set x = t.Cell(9, 9) : Out "cell 9,9:" & Err.Number : Err.Clear
Set x = w.Documents(5) : Out "missing document:" & Err.Number : Err.Clear
Set x = w.Documents.Open(target & ".missing.docx") : Out "open missing:" & Err.Number : Err.Clear
t.Cell(1, 1).Range = Null : Out "null:" & Err.Number : Err.Clear
On Error GoTo 0
d.SaveAs target
Out (d.Name = fso.GetFileName(target)) & "|" & (d.FullName = target) & "|" & (d.Path = fso.GetParentFolderName(target)) & "|" & d.Saved
d.Close
Out w.Documents.Count

' the saved document read again
Set d = w.Documents.Open(target)
Out d.Tables.Count & "|" & d.Tables(1).Rows.Count & "|" & d.Bookmarks.Count & "|" & Show(d.Content.Text)
For Each x In d.Tables(1).Rows(2).Cells : Out "cell:" & Show(x.Range.Text) : Next
For Each x In d.Bookmarks : Out "bookmark:" & x.Name : Next
d.Close False

' a new document
Set d = w.Documents.Add()
Out TypeName(d) & "|" & Left(d.Name, 8) & "|" & (d.FullName = d.Name) & "|[" & d.Path & "]|" & Show(d.Content.Text) & "|" & d.Saved
d.Content.Text = "Zeile 1" & vbCr & "Zeile 2" & vbTab & "Ende"
Out Show(d.Content) & "|" & d.Saved
Set t = d.Tables.Add(d.Content, 2, 2)
t.Cell(1, 1).Range.Text = "a"
t.Cell(2, 2).Range.Text = "d"
Out d.Tables.Count & "|" & Show(d.Content.Text)
d.Bookmarks.Add "ganz"
d.Bookmarks.Add "zelle", t.Cell(2, 2).Range
Out d.Bookmarks.Count & "|" & Show(d.Bookmarks("zelle").Range.Text) & "|" & d.Bookmarks(1).Name
d.Close False
Out w.Documents.Count
w.Quit
fso.DeleteFile target
