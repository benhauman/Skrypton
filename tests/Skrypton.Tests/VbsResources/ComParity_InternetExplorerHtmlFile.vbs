' Parity script: the expected output (ComParity_InternetExplorerHtmlFile.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' The clipboard contents are not compared: an htmlfile created by cscript accepts setData without changing the user's clipboard.
Dim h : Set h = CreateObject("htmlfile")
Out TypeName(h) & "|" & VarType(h) & "|" & h
Out TypeName(h.parentWindow) & "|" & VarType(h.parentWindow)
Dim cd : Set cd = h.parentWindow.clipboardData
Out TypeName(cd) & "|" & VarType(cd)
Dim r : r = cd.setData("text", "skrypton-clip")
Out TypeName(r) & "|" & r
Out TypeName(cd.getData("Text"))
Out TypeName(cd.clearData("text")) & "|" & TypeName(cd.clearData())
Dim ie : Set ie = CreateObject("InternetExplorer.Application")
Out TypeName(ie) & "|" & VarType(ie) & "|" & ie & "|" & ie.Name
Out ie.Visible & "|" & ie.Busy & "|" & ie.ReadyState & "|[" & ie.LocationURL & "]|[" & ie.LocationName & "]"
Out ie.Silent & "|" & ie.ToolBar & "|" & ie.StatusBar & "|" & ie.MenuBar & "|" & ie.AddressBar & "|" & ie.FullScreen & "|" & ie.Resizable & "|" & ie.Offline
Out TypeName(ie.Width) & "|" & TypeName(ie.Height) & "|" & TypeName(ie.Left) & "|" & TypeName(ie.Top) & "|" & TypeName(ie.HWND)
ie.Silent = True : ie.ToolBar = 0 : ie.StatusBar = True
Out ie.Silent & "|" & ie.ToolBar & "|" & ie.StatusBar
ie.Navigate "about:blank"
Do While ie.Busy Or ie.ReadyState <> 4 : WaitALittle : Loop
Out ie.ReadyState & "|" & ie.LocationURL & "|" & ie.Busy
ie.Quit
Sub WaitALittle()
  Dim t : t = Timer
  Do While Timer - t < 0.1 : Loop
End Sub
