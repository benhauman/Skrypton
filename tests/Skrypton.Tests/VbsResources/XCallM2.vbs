sub TestFso
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fso, BodyText, f
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile("C:\TRUMPF\helpLine\IntermediateReply.html", ForReading)
	BodyText = f.ReadAll
end sub