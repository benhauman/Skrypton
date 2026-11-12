' WScript.Echo (issue #119): the test runs this script with cscript and the command line tool and compares the output - Booleans are -1 / 0,
' Null is null with Echo, not with StdOut; Doubles have 15 and Singles 7 significant digits
WScript.Echo True
WScript.Echo False
WScript.Echo True, False, "x", 1, Null, Empty
Dim b : b = (1 = 1)
WScript.Echo b
WScript.Echo CBool(0)
WScript.StdOut.WriteLine True
WScript.StdOut.Write False
WScript.StdOut.WriteLine
WScript.Echo "" & True
WScript.Echo 1.5, #2026-11-12#, TimeSerial(10, 30, 0), 1 / 3, CSng(1 / 3), CCur(2.5), CByte(7), Exp(1)
WScript.Echo Empty
WScript.StdOut.WriteLine 1 / 3
WScript.StdOut.WriteLine CSng(1 / 3)
