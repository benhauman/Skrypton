SUB ButtonShowWebsite_Click()

'If hlObj.HasContent("PersonBilling.CostCenter_CA",0,0) = 0 Then
'	model.MsgBox "Bitte zuerst eine Kostenstelle erfassen!"
'	model.CurrentCommand.abort"OnSave"
'End if

If hlObj.IsNew = 1 Then

	dim sServer, sConn, oConn, sDatabaseName, sUser, sPassword 
	sDatabaseName="HLData" 
	sServer="MSSQLB" 
	sUser="helplinedata" 
	sPassword="helplinedata" 
	sConn="provider=sqloledb;data source=" & sServer & ";initial catalog=" & sDatabaseName
	Set oConn = CreateObject("ADODB.Connection") 
	oConn.Open sConn, sUser, sPassword 


	Const adCmdStoredProc = 4
	Const adInteger = 3
	Const adVarWChar = 202
	Const adParamInput = &H0001
	Const adParamOutput = &H0002
	Const adParamReturnValue = &H0004 
	Dim parmname,parmval, FirstCharName, xvIdentifier, group

	FirstCharName = LEFT(hlObj.GetValue("PersonGeneral.Name",0,0,0,0), 1)

	'SB Code ermitteln
	parmname="runScript" 
	Set adoSQLCmdParam = CreateObject("ADODB.Command")
	With adoSQLCmdParam
	Set .ActiveConnection = oConn
	.CommandText = "CreateNewSBCode"
	.CommandType = adCmdStoredProc
	.Parameters.Append .CreateParameter("RETURN_VALUEx", adInteger, adParamReturnValue )
	.Parameters.Append .CreateParameter("@FirstCharName", _
	adVarWChar, adParamInput,1,FirstCharName) 
	.Parameters.Append .CreateParameter("@NewSBCode", _
	adVarWChar, adParamOutput,10)
	.Execute
	parmval = .Parameters(2).Value
	End With

	hlObj.SetValue "PersonInformation.SBCode",0,0,0,parmval

	group = hlObj.GetValue("PersonGeneral.Group",0,0,0,0)

	If group = "GroupMainova" Or group = "GroupHolding" Then
		xvIdentifier = "X"
	Else
		xvIdentifier = "V"
	End If

	'X/V Personalnummer ermitteln
	Set adoSQLCmdParam2 = CreateObject("ADODB.Command")
	With adoSQLCmdParam2
	Set .ActiveConnection = oConn
	.CommandText = "CreateNewPersonalID"
	.CommandType = adCmdStoredProc
	.Parameters.Append .CreateParameter("RETURN_VALUEy", _
	adInteger, adParamReturnValue )
	.Parameters.Append .CreateParameter("@TypeCode", _
	adVarWChar, adParamInput,1,xvIdentifier) 
	.Parameters.Append .CreateParameter("@NewPersonalID", _
	adVarWChar, adParamOutput,10)
	.Execute
	parmval = .Parameters(2).Value
	End With

	hlObj.SetValue "PersonGeneral.PersonalID",0,0,0,parmval

	oConn.Close 
	Set oConn = Nothing

End If 
END SUB
