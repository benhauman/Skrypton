Public Sub ButtonShowWebsite_Click()

  'If hlObj.HasContent("PersonBilling.CostCenter_CA",0,0) = 0 Then
  '	model.MsgBox "Bitte zuerst eine Kostenstelle erfassen!"
  '	model.CurrentCommand.abort"OnSave"
  'End if

  IF hlObj.IsNew = 1 THEN

    Dim sServer, sConn, oConn, sDatabaseName, sUser, sPassword
    sDatabaseName = "HLData"
    sServer = "MSSQLB"
    sUser = "helplinedata"
    sPassword = "helplinedata"
    sConn = "provider=sqloledb;data source=" & sServer & ";initial catalog=" & sDatabaseName
    Set oConn = CreateObject("ADODB.Connection")
    oConn.Open sConn, sUser, sPassword


    Const adCmdStoredProc = 4

    Const adInteger = 3

    Const adVarWChar = 202

    Const adParamInput = 1

    Const adParamOutput = 2

    Const adParamReturnValue = 4

    Dim parmname, parmval, FirstCharName, xvIdentifier, group

    FirstCharName = LEFT(hlObj.GetValue("PersonGeneral.Name", 0, 0, 0, 0), 1)

    'SB Code ermitteln
    parmname = "runScript"
    Set adoSQLCmdParam = CreateObject("ADODB.Command")
    WITH adoSQLCmdParam
      Set .ActiveConnection = oConn
      .CommandText = "CreateNewSBCode"
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("@FirstCharName", adVarWChar, adParamInput, 1, FirstCharName)
      .Parameters.Append .CreateParameter("@NewSBCode", adVarWChar, adParamOutput, 10)
      .Execute
      parmval = .Parameters(2).Value
    END WITH

    hlObj.SetValue "PersonInformation.SBCode", 0, 0, 0, parmval

    group = hlObj.GetValue("PersonGeneral.Group", 0, 0, 0, 0)

    IF group = "GroupMainova" Or group = "GroupHolding" THEN
      xvIdentifier = "X"
    ELSE
      xvIdentifier = "V"
    END IF

    'X/V Personalnummer ermitteln
    Set adoSQLCmdParam2 = CreateObject("ADODB.Command")
    WITH adoSQLCmdParam2
      Set .ActiveConnection = oConn
      .CommandText = "CreateNewPersonalID"
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("@TypeCode", adVarWChar, adParamInput, 1, xvIdentifier)
      .Parameters.Append .CreateParameter("@NewPersonalID", adVarWChar, adParamOutput, 10)
      .Execute
      parmval = .Parameters(2).Value
    END WITH

    hlObj.SetValue "PersonGeneral.PersonalID", 0, 0, 0, parmval

    oConn.Close
    Set oConn = Nothing

  END IF
End Sub
