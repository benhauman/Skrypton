Dim serv
Set serv = Nothing
if not serv is nothing then
	If serv.enabled ( 7 ) = True Then
		serv.Enabled ( 8 ) = False
	Else
		serv.Enabled ( 9) = True 
	End If
end if
