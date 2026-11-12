' Parity script: the expected output (ComParity_FileSystemObjectEncodings.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' FileSystemObject text files (issue #163): TristateFalse reads and writes in the ANSI code page, TristateTrue UTF-16 LE (with its byte order mark),
' TristateUseDefault reads UTF-16 LE after its byte order mark and ANSI otherwise; File.OpenAsTextStream as OpenTextFile.
' (An ANSI file read as UTF-16 is left out: scrrun keeps invalid code units, .NET makes them U+FFFD.)
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim dir : dir = fso.GetSpecialFolder(2)
Function Bytes(p)
    Dim s : Set s = CreateObject("ADODB.Stream") : s.Type = 1 : s.Open : s.LoadFromFile p
    Dim b : b = s.Read : s.Close
    Dim i, r : For i = 1 To LenB(b) : r = r & Right("0" & Hex(AscB(MidB(b, i, 1))), 2) & " " : Next
    Bytes = Trim(r)
End Function
Sub WriteText(p, charset, text)
    Dim s : Set s = CreateObject("ADODB.Stream") : s.Type = 2 : s.Charset = charset : s.Open : s.WriteText text : s.SaveToFile p, 2 : s.Close
End Sub
Dim t : t = "Gr" & ChrW(252) & ChrW(223) & "e"
Dim pa : pa = fso.BuildPath(dir, "sk163-ansi.txt") : WriteText pa, "iso-8859-1", t
Dim pu8 : pu8 = fso.BuildPath(dir, "sk163-utf8.txt") : WriteText pu8, "utf-8", t
Dim pu16 : pu16 = fso.BuildPath(dir, "sk163-utf16.txt") : WriteText pu16, "unicode", t
Out "files: " & Bytes(pa) & " | " & Bytes(pu8) & " | " & Bytes(pu16)
Dim f, fmt
For Each fmt In Array(0, -1, -2)
    If fmt <> -1 Then Set f = fso.OpenTextFile(pa, 1, False, fmt) : Out "ansi fmt " & fmt & ": " & Escape(f.ReadAll) : f.Close
    Set f = fso.OpenTextFile(pu8, 1, False, fmt) : Out "utf8bom fmt " & fmt & ": " & Escape(f.ReadAll) : f.Close
    Set f = fso.OpenTextFile(pu16, 1, False, fmt) : Out "utf16 fmt " & fmt & ": " & Escape(f.ReadAll) : f.Close
Next
Set f = fso.OpenTextFile(pa, 1) : Out "ansi default: " & Escape(f.ReadAll) : f.Close
Set f = fso.OpenTextFile(pa, 1) : Out "ansi ReadLine: " & Escape(f.ReadLine) : f.Close
Set f = fso.GetFile(pa).OpenAsTextStream(1) : Out "ansi OpenAsTextStream: " & Escape(f.ReadAll) : f.Close
Set f = fso.GetFile(pu16).OpenAsTextStream(1, -1) : Out "utf16 OpenAsTextStream -1: " & Escape(f.ReadAll) : f.Close
Dim pw : pw = fso.BuildPath(dir, "sk163-write.txt")
Set f = fso.CreateTextFile(pw, True) : f.Write t : f.Close : Out "CreateTextFile ansi: " & Bytes(pw)
Set f = fso.CreateTextFile(pw, True, True) : f.Write "Gr" & ChrW(252) : f.Close : Out "CreateTextFile unicode: " & Bytes(pw)
Set f = fso.OpenTextFile(pw, 2, True, 0) : f.Write "Gr" & ChrW(252) & ChrW(8364) : f.Close : Out "OpenTextFile write 0: " & Bytes(pw)
Set f = fso.OpenTextFile(pw, 2, True, -1) : f.Write "Gr" & ChrW(252) : f.Close : Out "OpenTextFile write -1: " & Bytes(pw)
Set f = fso.OpenTextFile(pw, 8, True, -1) : f.Write "!" : f.Close : Out "OpenTextFile append -1: " & Bytes(pw)
Set f = fso.OpenTextFile(pw, 2, True, -2) : f.Write "Gr" & ChrW(252) : f.Close : Out "OpenTextFile write -2: " & Bytes(pw)
Set f = fso.OpenTextFile(pw, 8, True, 0) : f.Write ChrW(223) : f.Close : Out "OpenTextFile append 0: " & Bytes(pw)
Set f = fso.GetFile(pw).OpenAsTextStream(2, -1) : f.Write "x" : f.Close : Out "OpenAsTextStream write -1: " & Bytes(pw)
fso.DeleteFile pa : fso.DeleteFile pu8 : fso.DeleteFile pu16 : fso.DeleteFile pw
