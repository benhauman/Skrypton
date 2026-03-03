SUB Button1_Click()

Dim URL 
URL = hlObj.GetValue("vRealize.LansweeperURL",0,0,0,0) 

Dim wshShell, oExec 
Set wshShell = CreateObject("WScript.Shell") 
    wshShell.run URL

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

intProcessId = ""
For Each Process In Processes
    If StrComp(Process.Name, "iexplore.exe", vbTextCompare) = 0 Then
        intProcessId = Process.ProcessId
        Exit For
    End If
Next

If Len(intProcessId) > 0 Then
    With CreateObject("WScript.Shell")
        .AppActivate intProcessId

    End With
End If
END SUB
