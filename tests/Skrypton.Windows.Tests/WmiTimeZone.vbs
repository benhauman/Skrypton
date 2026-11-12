' Win32_TimeZone (issue #101): the test runs this script with cscript and the command line tool and compares the output (it depends on the time zone of the machine)
Dim svc : Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Dim col : Set col = svc.ExecQuery("Select * from Win32_TimeZone")
WScript.Echo "Count=" & col.Count & " " & TypeName(col)
Dim tz, n
For Each tz In col
    n = n + 1
    WScript.Echo TypeName(tz)
    WScript.Echo tz.Bias & "|" & VarType(tz.Bias) & "|" & tz.Caption & "|" & tz.Description & "|" & VarType(tz.SettingID)
    WScript.Echo tz.StandardName & "|" & tz.StandardBias & "|" & tz.StandardYear & "|" & tz.StandardMonth & "|" & tz.StandardDay & "|" & tz.StandardDayOfWeek & "|" & VarType(tz.StandardDayOfWeek) & "|" & tz.StandardHour & "|" & tz.StandardMinute & "|" & tz.StandardSecond & "|" & tz.StandardMillisecond
    WScript.Echo tz.DaylightName & "|" & tz.DaylightBias & "|" & tz.DaylightYear & "|" & tz.DaylightMonth & "|" & tz.DaylightDay & "|" & tz.DaylightDayOfWeek & "|" & VarType(tz.DaylightDayOfWeek) & "|" & tz.DaylightHour & "|" & tz.DaylightMinute & "|" & tz.DaylightSecond & "|" & tz.DaylightMillisecond
    Exit For
Next
WScript.Echo "n=" & n & " " & svc.InstancesOf("Win32_TimeZone").Count
Dim offsetMin
For Each tz In svc.ExecQuery("Select * from Win32_TimeZone")
    offsetMin = -(tz.Bias)
    Exit For
Next
WScript.Echo "offsetMin=" & offsetMin & " IsNumeric=" & IsNumeric(offsetMin) & " UTC=" & DateAdd("n", offsetMin, CDate("2026-11-12 10:00:00"))
