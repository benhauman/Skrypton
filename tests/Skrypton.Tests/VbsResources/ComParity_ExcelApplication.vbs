' Parity script: the expected output (ComParity_ExcelApplication.expected.txt) was captured with Windows cscript.exe and Microsoft Excel.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' SKRYPTON_PARITY_XLSX is a copy of ComParity_ExcelApplication.template.xlsx (created with Excel: sheet "Daten" with a bold header row,
' values, a "0.00" number format, dates and the formula =SUM(B2:B3)).
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh : Set sh = CreateObject("WScript.Shell")
Dim x, s, v
Dim template : template = sh.ExpandEnvironmentStrings("%SKRYPTON_PARITY_XLSX%")
Dim target : target = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton_parity_" & fso.GetTempName() & ".xlsx")

Dim xl : Set xl = CreateObject("Excel.Application")
Out TypeName(xl) & "|" & VarType(xl) & "|" & xl & "|" & xl.Name & "|" & xl.Visible & "|" & xl.DisplayAlerts
xl.DisplayAlerts = False
Out xl.DisplayAlerts & "|" & TypeName(xl.Workbooks) & "|" & xl.Workbooks.Count & "|" & TypeName(xl.ActiveWorkbook)

' a template read and filled in
Dim wb : Set wb = xl.Workbooks.Open(template)
Out TypeName(wb) & "|" & VarType(wb) & "|" & (wb.Name = fso.GetFileName(template)) & "|" & xl.Workbooks.Count & "|" & (xl.ActiveWorkbook.Name = wb.Name)
Out TypeName(wb.Worksheets) & "|" & TypeName(wb.Sheets) & "|" & wb.Worksheets.Count & "|" & wb.Worksheets(1).Name & "|" & wb.Worksheets("DATEN").Name & "|" & xl.ActiveSheet.Name
Dim ws : Set ws = wb.Worksheets("Daten")
Out TypeName(ws) & "|" & VarType(ws) & "|" & ws.Index & "|" & TypeName(ws.Cells(1, 1)) & "|" & VarType(ws.Cells(1, 1)) & "|" & ws.Cells(1, 1)
Out ws.Cells(1, 1).Value & "|" & ws.Range("B1").Value & "|" & ws.Cells(1, "C").Value & "|" & ws.Range("D1")
Out TypeName(ws.Cells(2, 2).Value) & ":" & ws.Cells(2, 2).Value & "|" & TypeName(ws.Cells(2, 3).Value) & ":" & ws.Cells(2, 3).Value & "|" & TypeName(ws.Cells(2, 4).Value) & ":" & Year(ws.Cells(2, 4).Value) & "-" & Month(ws.Cells(2, 4).Value) & "-" & Day(ws.Cells(2, 4).Value)
Out ws.Cells(4, 2).Formula & "|" & TypeName(ws.Cells(4, 2).Value) & ":" & ws.Cells(4, 2).Value & "|" & ws.Cells(2, 2).Formula & "|" & ws.Cells(2, 4).Formula
Out ws.Cells(1, 1).NumberFormat & "|" & ws.Cells(2, 3).NumberFormat & "|" & ws.Cells(2, 4).NumberFormat & "|" & ws.Cells(1, 1).Font.Bold & "|" & ws.Cells(2, 1).Font.Bold & "|" & TypeName(ws.Cells(1, 1).Font)
Out TypeName(ws.Cells(9, 9).Value) & "|" & VarType(ws.Cells(9, 9)) & "|[" & ws.Cells(9, 9).Text & "]|" & ws.Cells(2, 2).Text & "|" & ws.Cells(2, 1).Text
Out ws.UsedRange.Rows.Count & "x" & ws.UsedRange.Columns.Count & "|" & ws.UsedRange.Address & "|" & ws.Range("B2").Row & "," & ws.Range("B2").Column & "|" & ws.Cells(2, 2).Address & "|" & ws.Range("b2").Address(False, False)
Out ws.Range("A1:C2").Rows.Count & "|" & ws.Range("A1:C2").Columns.Count & "|" & ws.Range("A1:C2").Count & "|" & ws.Range("A1:C2").Address & "|" & ws.Range(ws.Cells(1, 1), ws.Cells(3, 2)).Address
v = ws.Range("A1:B2").Value
Out VarType(v) & "|" & LBound(v, 1) & "-" & UBound(v, 1) & "|" & LBound(v, 2) & "-" & UBound(v, 2) & "|" & v(1, 1) & "|" & v(2, 2)
ws.Cells(5, 1).Value = "Bolzen"
ws.Cells(5, 2) = 25
ws.Range("C5").Value = "0.5"
ws.Cells(5, 4).Value = DateSerial(2026, 11, 11)
ws.Cells(6, 1).Value = "007"
ws.Cells(6, 2).Value = "TRUE"
ws.Cells(6, 3).Value = True
ws.Cells(6, 4).Value = "Text"
Out TypeName(ws.Cells(5, 2).Value) & "|" & TypeName(ws.Cells(5, 3).Value) & "|" & TypeName(ws.Cells(5, 4).Value) & "|" & ws.Cells(5, 4).NumberFormat & "|" & TypeName(ws.Cells(6, 1).Value) & ":" & ws.Cells(6, 1).Value & "|" & TypeName(ws.Cells(6, 2).Value) & "|" & ws.Cells(6, 3).Text & "|" & ws.Cells(6, 4).Value
ws.Cells(4, 2).Formula = "=SUM(B2:B5)"
ws.Range("A6:D6").Font.Bold = True
ws.Range("B5").NumberFormat = "0.00"
Out ws.Cells(4, 2).Formula & "|" & ws.Cells(6, 4).Font.Bold & "|" & ws.Cells(5, 2).NumberFormat & "|" & ws.Cells(5, 2).Value & "|" & wb.Saved
ws.Range("A7:B7").Value = "x"
ws.Cells(6, 4).Value = Empty
Out ws.Cells(7, 1).Value & ws.Cells(7, 2).Value & "|" & TypeName(ws.Cells(6, 4).Value) & "|" & ws.UsedRange.Address
On Error Resume Next
Set s = wb.Worksheets("Missing") : Out "missing sheet:" & Err.Number : Err.Clear
Set s = wb.Worksheets(9) : Out "sheet 9:" & Err.Number : Err.Clear
Set s = xl.Workbooks(5) : Out "workbook 5:" & Err.Number : Err.Clear
Set s = xl.Workbooks.Open(target & ".missing.xlsx") : Out "open missing:" & Err.Number : Err.Clear
Set s = ws.Range("not an address") : Out "bad range:" & Err.Number : Err.Clear
On Error GoTo 0
Dim ws2 : Set ws2 = wb.Worksheets.Add()
Out wb.Worksheets.Count & "|" & ws2.Name & "|" & wb.Worksheets(1).Name & "|" & xl.ActiveSheet.Name
ws2.Name = "Neu"
ws2.Cells(1, 1).Value = "auf Neu"
wb.SaveAs target
Out (wb.Name = fso.GetFileName(target)) & "|" & (wb.FullName = target) & "|" & (wb.Path = fso.GetParentFolderName(target)) & "|" & wb.Saved
wb.Close False
Out xl.Workbooks.Count & "|" & TypeName(xl.ActiveWorkbook)

' the saved workbook read again
Set wb = xl.Workbooks.Open(target)
Out wb.Worksheets.Count & "|" & wb.Worksheets(1).Name & "|" & wb.Worksheets(2).Name
Set ws = wb.Worksheets("Daten")
Out ws.Cells(5, 1).Value & "|" & ws.Cells(5, 2).Value & "|" & TypeName(ws.Cells(5, 4).Value) & "|" & ws.Cells(6, 1).Value & "|" & ws.Cells(4, 2).Formula & "|" & ws.Cells(6, 1).Font.Bold & "|" & ws.Cells(5, 2).NumberFormat & "|" & wb.Worksheets("Neu").Cells(1, 1).Value
For Each s In wb.Worksheets : Out "sheet:" & s.Name : Next
wb.Close False

' a new workbook
Set wb = xl.Workbooks.Add()
Out TypeName(wb) & "|" & Left(wb.Name, 4) & "|" & (wb.FullName = wb.Name) & "|[" & wb.Path & "]|" & wb.Worksheets.Count & "|" & wb.Worksheets(1).Name & "|" & wb.Saved
xl.Cells(1, 1).Value = "via Application"
xl.Range("B1").Value = 2
Out wb.Worksheets(1).Cells(1, 1).Value & "|" & xl.ActiveSheet.Range("B1").Value & "|" & TypeName(xl.Columns("A:B").AutoFit) & "|" & xl.ActiveSheet.UsedRange.Address & "|" & xl.ActiveSheet.UsedRange.Rows.Count & "x" & xl.ActiveSheet.UsedRange.Columns.Count
wb.Close False
xl.Quit
fso.DeleteFile target
