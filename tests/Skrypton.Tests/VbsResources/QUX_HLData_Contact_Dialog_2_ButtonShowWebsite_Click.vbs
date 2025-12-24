Dim objShell
set objShell = CreateObject("Shell.Application")
If Trim(TextBoxWebsite.Text) <> "" Then
	objShell.ShellExecute TextBoxWebsite.Text
End If