' WshShell.RegRead (issue #122): the test runs this script with cscript and the command line tool and compares the output (the values depend on the machine)
Dim sh : Set sh = CreateObject("WScript.Shell")
Sub Show(name)
    On Error Resume Next
    Dim v : v = sh.RegRead(name)
    If Err.Number <> 0 Then
        WScript.Echo name & " => ERR " & Err.Number & " [" & Err.Description & "] " & Err.Source
        Err.Clear
        Exit Sub
    End If
    Dim s
    If IsArray(v) Then
        s = "array(" & (UBound(v) + 1) & ") " & TypeName(v(0))
    Else
        s = Len(CStr(v)) & " chars"
    End If
    WScript.Echo name & " => " & VarType(v) & " " & TypeName(v) & " " & s
End Sub
Show "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
Show "HKLM\System\CurrentControlSet\Control\TimeZoneInformation\TimeZoneKeyName"
Show "HKCU\Environment\TEMP"
Show "HKLM\SYSTEM\CurrentControlSet\Control\ServiceGroupOrder\List"
Show "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"
Show "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\InstallTime"
Show "HKCR\.txt\"
Show "HKEY_CLASSES_ROOT\.txt\"
Show "HKU\.DEFAULT\Environment\TEMP"
Show "HKCC\System\CurrentControlSet\Control\Print\Printers\"
Show "HKLM\SOFTWARE\NoSuchKeySkrypton\Value"
Show "HKLM\SOFTWARE\Microsoft\NoSuchValueSkrypton"
Show "HKLM\SOFTWARE\NoSuchKeySkrypton\"
Show "HKXX\Software\x"
Show "HKCU\Environment\temp"
WScript.Echo sh.RegRead("HKCU\Environment\TEMP")
Show "HKEY_CURRENT_CONFIG\Software\"
Show "HKLM\SOFTWARE\Microsoft\"
Show "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber"
Show "hklm\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber"
Show "HKLM"
Show "HKLM\"
Show ""
Show "HKEY_USERS\.DEFAULT\Environment\TEMP"
