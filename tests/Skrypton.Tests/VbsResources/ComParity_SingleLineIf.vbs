' Parity script: the expected output (ComParity_SingleLineIf.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' A single-line If may end with End If - also after ":" and with statements after it on the same line (issue #97)
Dim c, r
c = True
r = Empty
If c Then r = "a" Else r = "b" End If
Out "[If c Then r = ""a"" Else r = ""b"" End If] " & r
r = Empty
If True Then r = "t" End If
Out "[If True Then r = ""t"" End If] " & r
r = Empty
If True Then r = "u" : r = r & "v" End If
Out "[If True Then r = ""u"" : r = r & ""v"" End If] " & r
r = Empty
If False Then r = "u" Else r = "w" : r = r & "z" End If
Out "[If False Then r = ""u"" Else r = ""w"" : r = r & ""z"" End If] " & r
r = Empty
Function T(n) : If n Then T = "yes" Else T = "no" End If : End Function : r = T(True) & T(False)
Out "[Function T(n) : If n Then T = ""yes"" Else T = ""no"" End If : End Function : r = T(True) & T(False)] " & r
r = Empty
If True Then r = "e" End If : r = r & "!"
Out "[If True Then r = ""e"" End If : r = r & ""!""] " & r
r = Empty
If True Then r = 1 : End If
Out "[If True Then r = 1 : End If] " & r
r = Empty
If True Then r = 1 Else r = 2 : End If
Out "[If True Then r = 1 Else r = 2 : End If] " & r
r = Empty
If False Then r = 1 Else r = 2 : End If : r = r & "!"
Out "[If False Then r = 1 Else r = 2 : End If : r = r & ""!""] " & r
r = Empty
Sub S : If True Then r = "s" End If : End Sub : S
Out "[Sub S : If True Then r = ""s"" End If : End Sub : S] " & r
r = Empty
Sub S2 : If True Then r = "s2" : End If : End Sub : S2
Out "[Sub S2 : If True Then r = ""s2"" : End If : End Sub : S2] " & r
r = Empty
If True Then If True Then r = "nn" End If
Out "[If True Then If True Then r = ""nn"" End If] " & r
r = Empty
If c Then r = "q" ' comment
Out "[If c Then r = ""q"" ' comment] " & r
r = Empty
If c Then r = "x" Else If False Then r = 1 Else r = "ee" End If : r = r & "!"
Out "[If c Then r = ""x"" Else If False Then r = 1 Else r = ""ee"" End If : r = r & ""!""] " & r
r = Empty
If False Then r = "x" Else If False Then r = 1 Else r = "ee" End If : r = r & "!"
Out "[If False Then r = ""x"" Else If False Then r = 1 Else r = ""ee"" End If : r = r & ""!""] " & r
r = Empty
If False Then r = "x" Else If True Then r = 1 End If : r = r & "!" End If : r = r & "?"
Out "[If False Then r = ""x"" Else If True Then r = 1 End If : r = r & ""!"" End If : r = r & ""?""] " & r
c = False
r = Empty
If c Then r = "a" Else r = "b" End If
Out "[c = False] " & r
