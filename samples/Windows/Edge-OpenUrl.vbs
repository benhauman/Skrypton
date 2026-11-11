' Demo: open a web page in Microsoft Edge, and use the browser object (InternetExplorer.Application, which Windows forwards to Edge).
'   Skrypton.Windows.Cli Edge-OpenUrl.vbs [url]           starts Edge with the page
'   Skrypton.Windows.Cli Edge-OpenUrl.vbs /dryrun [url]   only prints what it would do
Option Explicit
Dim dryRun, url, i, shell, command, browser
dryRun = False
url = "https://github.com/benhauman/Skrypton"
For i = 0 To WScript.Arguments.Count - 1
    If LCase(WScript.Arguments(i)) = "/dryrun" Then
        dryRun = True
    Else
        url = WScript.Arguments(i)
    End If
Next

Set shell = CreateObject("WScript.Shell")
command = "msedge.exe """ & url & """"
WScript.Echo "Command: " & command

Set browser = CreateObject("InternetExplorer.Application")
WScript.Echo "Browser object: " & TypeName(browser) & ", " & browser.Name & ", visible: " & browser.Visible & ", ready state: " & browser.ReadyState
browser.Quit

If dryRun Then
    WScript.Echo "Dry run: Edge is not started"
Else
    shell.Run command, 1, False
End If
