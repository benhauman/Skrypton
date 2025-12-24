Dim objShell
Set objShell = CreateObject("Shell.Application")
IF Trim(TextBoxWebsite.Text) <> "" THEN
  objShell.ShellExecute TextBoxWebsite.Text
END IF
