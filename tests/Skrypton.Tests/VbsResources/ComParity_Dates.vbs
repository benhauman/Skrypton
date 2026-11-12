' Parity script: the expected output (ComParity_Dates.expected.txt) was captured with Windows cscript.exe (en-US).
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' CDate of a text without a date is the time on 1899-12-30, not today (issue #124)
Dim t
For Each t In Array("10:30:00", "10:30", "10:30 PM", "22:15:05", "0:00", "12:00:00 AM", "2026-11-12 10:30", "11/12/2026", "11/12/2026 10:30:00 AM", "Nov 12, 2026")
    Out t & " => " & CDate(t) & "|" & Year(CDate(t)) & "|" & CDbl(CDate(t)) & "|" & IsDate(t)
Next
Dim same : same = (CDate("10:30:00") = TimeSerial(10, 30, 0)) : Out "" & same
Out DateAdd("n", 30, "10:30")
Out TimeValue("10:30") & "|" & DateValue("11/12/2026")
