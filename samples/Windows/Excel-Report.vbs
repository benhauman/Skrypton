' Demo: create an Excel report (bold header, values, number formats, a sum formula), save it, open it again and read it.
'   Skrypton.Windows.Cli Excel-Report.vbs   (Excel does not need to be installed)
Option Explicit
Dim fso, file, excel, workbook, sheet, row, items, item
Set fso = CreateObject("Scripting.FileSystemObject")
file = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-demo-report.xlsx")
items = Array(Array("Screws", 100, 0.25), Array("Nuts", 250, 0.1), Array("Bolts", 40, 1.75))

Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)
sheet.Name = "Report"
sheet.Cells(1, 1).Value = "Article"
sheet.Cells(1, 2).Value = "Quantity"
sheet.Cells(1, 3).Value = "Price"
sheet.Range("A1:C1").Font.Bold = True
row = 2
For Each item In items
    sheet.Cells(row, 1).Value = item(0)
    sheet.Cells(row, 2).Value = item(1)
    sheet.Cells(row, 3).Value = item(2)
    row = row + 1
Next
sheet.Range("C2:C4").NumberFormat = "0.00"
sheet.Cells(row, 1).Value = "Total"
sheet.Cells(row, 2).Formula = "=SUM(B2:B4)"
sheet.Columns("A:C").AutoFit
workbook.SaveAs file
workbook.Close False
WScript.Echo "Saved: " & fso.GetFileName(file)

Set workbook = excel.Workbooks.Open(file)
Set sheet = workbook.Worksheets("Report")
WScript.Echo "Sheet: " & sheet.Name & ", used range " & sheet.UsedRange.Address & " (" & sheet.UsedRange.Rows.Count & " rows)"
For row = 1 To 4
    WScript.Echo "  " & sheet.Cells(row, 1).Value & " | " & sheet.Cells(row, 2).Value & " | " & sheet.Cells(row, 3).Value
Next
WScript.Echo "  " & sheet.Cells(5, 1).Value & " | " & sheet.Cells(5, 2).Formula
WScript.Echo "Header bold: " & sheet.Cells(1, 1).Font.Bold & ", price format: " & sheet.Cells(2, 3).NumberFormat
workbook.Close False
excel.Quit
fso.DeleteFile file
