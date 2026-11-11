' Demo: an inventory of files - Scripting.FileSystemObject, Scripting.Dictionary and VBScript.RegExp.
'   Skrypton.Windows.Cli FileSystem-Inventory.vbs
Option Explicit
Dim fso, folder, names, name, file, counts, extension, regex, matches, keys, i, j, swap
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-demo-inventory")
If fso.FolderExists(folder) Then fso.DeleteFolder folder, True
fso.CreateFolder folder
names = Array("report-2026-01.txt", "report-2026-02.txt", "data.csv", "invoice-17.pdf", "notes.txt")
For Each name In names
    fso.CreateTextFile(fso.BuildPath(folder, name), True).Close
Next

Set counts = CreateObject("Scripting.Dictionary")
Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "(\d{4})-(\d{2})"
For Each file In fso.GetFolder(folder).Files
    extension = LCase(fso.GetExtensionName(file.Name))
    counts(extension) = counts(extension) + 1
    Set matches = regex.Execute(file.Name)
    If matches.Count > 0 Then WScript.Echo "Monthly file: " & file.Name & " (year " & matches(0).SubMatches(0) & ", month " & matches(0).SubMatches(1) & ")"
Next

keys = counts.Keys
For i = 0 To UBound(keys) - 1
    For j = i + 1 To UBound(keys)
        If keys(j) < keys(i) Then
            swap = keys(i)
            keys(i) = keys(j)
            keys(j) = swap
        End If
    Next
Next
For Each extension In keys
    WScript.Echo "." & extension & ": " & counts(extension) & " file(s)"
Next
WScript.Echo "Files: " & fso.GetFolder(folder).Files.Count
fso.DeleteFolder folder, True
