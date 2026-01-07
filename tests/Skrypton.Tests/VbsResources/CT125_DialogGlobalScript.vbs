Function ConvertSize(Size) 

	'MsgBox "Converting Size for " & Size
	Size = CSng(Replace(Size,",",""))
	
	If Not VarType(Size) = vbSingle Then 
		ConvertSize = "SIZE INPUT ERROR"
		Exit Function
	End If
	
	Suffix = " B" 
	If Size >= 1024 Then suffix = " KB" 
	If Size >= 1048576 Then suffix = " MB" 
	If Size >= 1073741824 Then suffix = " GB" 
	If Size >= 1099511627776 Then suffix = " TB" 
	
	Select Case Suffix 
		Case " KB" Size = Round(Size / 1024, 2) 
		Case " MB" Size = Round(Size / 1048576, 2) 
		Case " GB" Size = Round(Size / 1073741824, 2) 
		Case " TB" Size = Round(Size / 1099511627776, 2) 
	End Select

	ConvertSize = Size & Suffix 
End Function
Function getNexthinkUser()
	getNexthinkUser = "myusr2"
End Function
Function getNexthinkBaseURL()
	getNexthinkBaseURL = ""
End Function
Function getNexthinkPassword()
	getNexthinkPassword = "mypwd2"
End Function