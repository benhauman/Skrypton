' Parity script: the expected output (ComParity_BuiltInFunctions.expected.txt) was captured with Windows cscript.exe (culture en-US).
' The test runs this script through Skrypton (Out writes a line, culture en-US) and compares the output line by line.
Dim v, a, i
Sub Show(name, value)
    If IsArray(value) Then
        Out name & ": array(" & LBound(value) & " to " & UBound(value) & ") " & Join(value, ",")
    ElseIf IsNull(value) Then
        Out name & ": Null"
    Else
        Out name & ": " & TypeName(value) & " " & value
    End If
End Sub
Sub ShowError(name)
    Out name & ": error " & Err.Number
    Err.Clear
End Sub

' Characters and strings
Show "ChrW(252)", AscW(ChrW(252)) & " " & Len(ChrW(252)) & " " & TypeName(ChrW(252))
Show "ChrW(65)", ChrW(65)
Show "AscW(ChrW(8364))", AscW(ChrW(8364))
Show "ChrW(-1) AscW", AscW(ChrW(-1))
Show "StrReverse", StrReverse("Skrypton")
Show "StrReverse empty", StrReverse("")
Show "LenB", LenB("abc")
a = Array("apple", "Banana", "cherry", "date")
Show "Filter an", Filter(a, "an")
Show "Filter an exclude", Filter(a, "an", False)
Show "Filter A text", Filter(a, "A", True, vbTextCompare)
Show "Filter A binary", Filter(a, "A", True, vbBinaryCompare)
Show "Filter none", Filter(a, "xyz")

' Numbers
Show "Fix(2.7)", Fix(2.7)
Show "Fix(-2.7)", Fix(-2.7)
Show "Fix(5)", Fix(5)
Show "Fix(""-3.5"")", Fix("-3.5")
Show "Fix(Null)", Fix(Null)
Show "Int(-2.7)", Int(-2.7)
Show "Sgn(-5)", Sgn(-5)
Show "Sgn(0)", Sgn(0)
Show "Sgn(2.5)", Sgn(2.5)
Show "Exp(1)", Exp(1)
Show "Exp(0)", Exp(0)
Show "Log(1)", Log(1)
Show "Log(10)", Log(10)
Show "Tan(0)", Tan(0)
Show "Tan(1)", Tan(1)
Show "Oct(8)", Oct(8)
Show "Oct(255)", Oct(255)
Show "Oct(-1)", Oct(-1)
Show "Oct(-32769)", Oct(-32769)
Show "Hex(255)", Hex(255)
Show "True Eqv False", True Eqv False
Show "True Eqv True", True Eqv True
Show "5 Eqv 3", 5 Eqv 3
Show "True Imp False", True Imp False
Show "False Imp True", False Imp True
Show "5 Imp 3", 5 Imp 3
Show "Null Eqv True", Null Eqv True
Show "False Imp Null", False Imp Null

' Formats (en-US)
Show "FormatNumber(1234.5)", FormatNumber(1234.5)
Show "FormatNumber(1234.5, 0)", FormatNumber(1234.5, 0)
Show "FormatNumber(-0.5, 2)", FormatNumber(-0.5, 2)
Show "FormatNumber(0.5, 2, 0)", FormatNumber(0.5, 2, 0)
Show "FormatNumber(-1234.5, 1, -1, -1)", FormatNumber(-1234.5, 1, -1, -1)
Show "FormatNumber(1234.5, 1, -1, 0, 0)", FormatNumber(1234.5, 1, -1, 0, 0)
Show "FormatCurrency(1234.5)", FormatCurrency(1234.5)
Show "FormatCurrency(-1234.5)", FormatCurrency(-1234.5)
Show "FormatCurrency(1234.567, 1)", FormatCurrency(1234.567, 1)
Show "FormatPercent(0.256)", FormatPercent(0.256)
Show "FormatPercent(0.256, 0)", FormatPercent(0.256, 0)
Show "FormatPercent(1.5, 1)", FormatPercent(1.5, 1)

' Dates
v = DateSerial(2026, 11, 11) + TimeSerial(13, 45, 30)
Show "DatePart yyyy", DatePart("yyyy", v)
Show "DatePart q", DatePart("q", v)
Show "DatePart m", DatePart("m", v)
Show "DatePart y", DatePart("y", v)
Show "DatePart d", DatePart("d", v)
Show "DatePart w", DatePart("w", v)
Show "DatePart w monday", DatePart("w", v, vbMonday)
Show "DatePart ww", DatePart("ww", v)
Show "DatePart ww first full week", DatePart("ww", DateSerial(2026, 1, 3), vbSunday, vbFirstFullWeek)
Show "DatePart h", DatePart("h", v)
Show "DatePart n", DatePart("n", v)
Show "DatePart s", DatePart("s", v)
Show "MonthName(1)", MonthName(1)
Show "MonthName(11, True)", MonthName(11, True)
Show "WeekdayName(1)", WeekdayName(1)

' Script engine
Show "ScriptEngine", ScriptEngine()
Show "ScriptEngineMajorVersion", ScriptEngineMajorVersion()
Show "ScriptEngineMinorVersion", ScriptEngineMinorVersion()
Show "GetLocale", GetLocale()

' Errors
On Error Resume Next
v = ChrW(70000) : ShowError "ChrW(70000)"
v = Log(0) : ShowError "Log(0)"
v = Sgn(Null) : ShowError "Sgn(Null)"
v = StrReverse(Null) : ShowError "StrReverse(Null)"
v = MonthName(13) : ShowError "MonthName(13)"
v = DatePart("x", Now) : ShowError "DatePart x"
v = Filter("abc", "a") : ShowError "Filter not an array"
On Error GoTo 0
