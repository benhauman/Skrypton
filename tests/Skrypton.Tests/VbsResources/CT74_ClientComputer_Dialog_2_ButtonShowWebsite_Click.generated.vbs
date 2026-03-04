Public Sub Button1_Click()

  Dim URL
  URL = hlObj.GetValue("vRealize.LansweeperURL", 0, 0, 0, 0)

  Dim wshShell, oExec
  Set wshShell = CreateObject("WScript.Shell")
  wshShell.run URL

  Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

  intProcessId = ""
  For Each Process In Processes
    IF StrComp(Process.Name, "iexplore.exe", vbTextCompare) = 0 THEN
      intProcessId = Process.ProcessId
      Exit For
    END IF
  Next

  IF Len(intProcessId) > 0 THEN
    WITH CreateObject("WScript.Shell")
      .AppActivate intProcessId

    END WITH
  END IF
End Sub
