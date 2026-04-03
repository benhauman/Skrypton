SUB Test
	Dim responsiblity, requesttype
	responsiblity = "Rs"
	requesttype = "Qq"
	
	Select Case responsiblity
		Case "EditorTypeConsulting"
			Select Case requesttype
				Case "RequestTypeService", "RequestTypeBug", "RequestTypeExtra", "RequestTypeInformation", "RequestTypeRegulier"
					hlObj.SetValue "IncidentAttribute.SLAControl",7,0,0,"SLAControlCons1"
				Case "RequestTypeDefect"
					hlObj.SetValue "IncidentAttribute.SLAControl",7,0,0,"SLAControlCons1"
			End Select
			
	End Select
	
	
END SUB