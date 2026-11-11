' Demo: create a Word letter (text, a table, bookmarks), save it, open it again, change a bookmark and read it.
'   Skrypton.Windows.Cli Word-Letter.vbs   (Word does not need to be installed)
Option Explicit
Dim fso, file, word, document, table, text
Set fso = CreateObject("Scripting.FileSystemObject")
file = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-demo-letter.docx")

Set word = CreateObject("Word.Application")
Set document = word.Documents.Add()
Set table = document.Tables.Add(document.Content, 2, 3)
table.Cell(1, 1).Range.Text = "Item"
table.Cell(1, 2).Range.Text = "Quantity"
table.Cell(1, 3).Range.Text = "Price"
table.Rows(2).Cells(1).Range = "Consulting"
table.Rows(2).Cells(2).Range = 8
table.Rows(2).Cells(3).Range = "960.00"
document.Bookmarks.Add "price", table.Cell(2, 3).Range
document.SaveAs file
document.Close
WScript.Echo "Saved: " & fso.GetFileName(file)

Set document = word.Documents.Open(file)
WScript.Echo "Tables: " & document.Tables.Count & ", rows: " & document.Tables(1).Rows.Count & ", bookmarks: " & document.Bookmarks.Count
document.Bookmarks("price").Range.Text = "1,200.00"
document.Tables(1).Rows.Add
document.Tables(1).Cell(3, 1).Range.Text = "Travel"
text = document.Content.Text
WScript.Echo "Content: " & Replace(Replace(text, Chr(13) & Chr(7), " | "), vbCr, " / ")
WScript.Echo "Price cell: " & Replace(document.Tables(1).Cell(2, 3).Range.Text, Chr(13) & Chr(7), "") & ", bookmark kept: " & document.Bookmarks.Exists("price")
document.Close False
word.Quit
fso.DeleteFile file
