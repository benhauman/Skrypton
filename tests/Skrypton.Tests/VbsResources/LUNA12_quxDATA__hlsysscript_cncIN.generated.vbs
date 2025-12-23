' StartITIL3
' Copyright 2001-2009 PMCS GmbH & Co. KG
' ü
' helpLine Connectivity Page 'CNCIN'
' Diese Datei beinhaltet die Logik, wie helpline Connectivity die eingehenden Mails bearbeitet
'-------------------------------------------------------------------------------------- script ---

Option Explicit

'--------------------------------------------------------------------------------------- sub 1 ---
Public Sub ProcessIn()
  ' main ü entry point for execution
  LogText "ProcessRequestMail start."

  Dim oMailRequest, oHLServer
  Dim adhocMail
  Dim autoReplyList
  Dim imKeywords
  Dim rfKeywords
  Dim cmKeywords
  Dim item

  Set oMailRequest = session("mailrequest")
  Set oHLServer = session("serverconnection")

  autoReplyList = Array("Out of Office:", "Abwesend:")
  rfKeywords = Array("[ServiceRequest]")
  imKeywords = Array("[Incident]")
  cmKeywords = Array("[RFC]")


  LogText "mail subject:" & oMailRequest.subject

  For Each item In autoReplyList
    IF (InStr(1, oMailRequest.Subject, item, 1) > 0) THEN
      session("processtext") = "Out of Office AutoReply"
      Exit Sub
    END IF
  Next

  oMailRequest.mailtype = - 2
  adhocMail = False
  adhocMail = IsAdhocMail(oMailRequest)

  '+++ Änderung für Workflow +++
  Dim sReportText
  Dim refNumber
  refNumber = ExtractRefNumber(oMailRequest.Subject)
  IF (Len(refNumber) > 0) THEN
    Dim caseToExtend
    Set caseToExtend = session.GetCaseByReferenceNumber(refNumber)
    LogText "RefNumber > 0"
    IF (session.IsBuiltinCase(caseToExtend)) THEN
      LogText "IsBuiltinCase"
      sReportText = extendCaseFromMail(oMailRequest, oCaseCfg, oHLServer, refNumber)
      Exit Sub
    ELSE
      LogText "NOT IsBuiltinCase"
      IF (session.CanExtendWorkflowCase(caseToExtend)) THEN
        LogText "CanExtend"
        sReportText = session.DoExtendWorkflowCase(caseToExtend)
        Exit Sub
      ELSE
        LogText "CanNotExtend"
      END IF
    END IF
  END IF
  'sReportText = session.NewWorkflowFromMail("AzureEvent")
  'If (IsWFEmail(oMailRequest.Subject, rfKeywords) = True) Then
  '	sReportText = session.NewWorkflowFromMail("RequestFulfillment")
  'Else
  '  If (IsWFEmail(oMailRequest.Subject, imKeywords) = True) Then
  ' 		sReportText = session.NewWorkflowFromMail("IncidentManagement")
  '	Else
  '   	If (IsWFEmail(oMailRequest.Subject, cmKeywords) = True) Then
  '		sReportText = session.NewWorkflowFromMail("ChangeManagement")
  '	Else
  '      If adhocMail = True Then
  '      	CreateAdhocCase oMailRequest
  '      Else
  sReportText = session.NewWorkflowFromMail("Request")
  '      End If
  '		End If
  '   End If
  ' End If

  LogText "ProcessRequestMail end."
End Sub

'--------------------------------------------------------------------------------------- sub 2 ---
Public Sub LogText(ByRef sText)
  'session("worker").trace sText
  session("processtext") = session("processtext") & sText & vbLf
End Sub

'--------------------------------------------------------------------------------------- sub 3 ---
Public Sub SetCaseAttributes(ByRef hlcase, ByRef mail)

  LogText "SetCaseAttributes"

  Dim oScripter
  Set oScripter = session("worker").CreateScriptEngine

  oScripter.AddObject "hlcase", hlcase
  oScripter.AddObject "mail", mail

  session("worker").ExecuteScript oScripter, session, "receive"

End Sub

Public Function IsAdhocMail(ByRef oMailRequest)

  '	Suche die Konfiguration für diesen Vorgangstypen

  Dim bRegisteredMailType
  bRegisteredMailType = False

  Dim oConfig
  Set oConfig = session("config")

  Dim oCaseCfgs, oCaseCfg, oCaseType
  Set oCaseCfgs = oConfig.GetGroup("CaseTypes")

  For Each oCaseType In oCaseCfgs.Groups
    IF (oCaseType.GetValue("type").data = oMailRequest.mailtype) THEN
      Set oCaseCfg = oCaseType
      oMailRequest.mailtype = oCaseCfg.GetValue("type").data
      bRegisteredMailType = True
      Exit For
    END IF
  Next

  IsAdhocMail = bRegisteredMailType
End Function

Public Sub CreateAdhocCase(ByRef oMailRequest)

  ' Suche die Objektdefinition anhand des Betreffs in der E-Mail

  Dim oSubjectValue
  For Each oSubjectValue In session("config").GetGroup("subject").values
    IF (InStr(1, oMailRequest.Subject, oSubjectValue.data, 1) > 0) THEN
      oMailRequest.mailtype = CLng(oSubjectValue.Name)
      Exit For
    END IF
  Next
  IF (oMailRequest.mailtype < 0) THEN
    session("processtext") = "unregistered mail subject"
    Exit Sub
  END IF
  LogText "MailRequestType:" & oMailRequest.mailtype
  sReportText = createCaseFromMail(oMailRequest, oCaseCfg, oHLServer)
End Sub

Public Sub SetSUAttributes(ByRef hlcase, ByRef mail)

  LogText "SetSUAttributes"

  Dim oScripter
  Set oScripter = session("worker").CreateScriptEngine

  oScripter.AddObject "hlcase", hlcase
  oScripter.AddObject "mail", mail

  session("worker").ExecuteScript oScripter, session, "extend"

End Sub


Public Sub AssociateSenderToCase(ByRef oMailRequest, ByRef oCaseCfg, ByRef oHLServer, ByRef oCase)


  ' Suche

  Dim sMailAttributeKey, sSearchConditionPersons, oPersons
  sMailAttributeKey = oCaseCfg.GetValue("MailAttributeKey").data
  sSearchConditionPersons = sMailAttributeKey & "= "" & oMailRequest.SenderMail & """

  LogText "SearchCondition = " & sSearchConditionPersons
  Set oPersons = oHLServer.Find_Persons(sSearchConditionPersons, 0)


  ' Baue eine Assoziation zwischen Vorgang und Anfrager

  IF oPersons.Count = 0 THEN
    Set oPersons = Nothing
    ' Keine Person mit der EmailAdresse gefunden !!!!
    ' Besser für Auswertung mit Berichten ist ein DummyPerson
    ' z.B. "email adresse unbekant" als Anfrager zu setzen

    ' Bitte zuerst in helpLine diese Dummy-Person anlegen !

    sSearchConditionPersons = "PersonGeneral.Name = "email adresse unbekannt""
    LogText "SearchCondition2 = " & sSearchConditionPersons
    Set oPersons = oHLServer.Find_Persons(sSearchConditionPersons, 0)
    IF oPersons.Count > 0 THEN
      oCase.AssociatePersons oPersons
    END IF
  ELSE
    oCase.AssociatePersons oPersons
  END IF

End Sub


'---------------------------------------------------------------------------------------- createCaseFromMail ---
Public Function CreateCaseFromMail(ByRef oMailRequest, ByRef oCaseCfg, ByRef oHLServer)

  LogText "createCaseFromMail"


  '	Erzeuge einen Vorgang


  Dim sCaseType, oCase, oHLCase
  sCaseType = oCaseCfg.GetValue("CaseType").data
  Set oCase = oHLServer.CreateCase(sCaseType)
  Set oHLCase = oCase.GetHLObject

  LogText "case-id:" & CStr(oHLCase.GetID)

  AssociateSenderToCase oMailRequest, oCaseCfg, oHLServer, oCase

  ' Setze die Attribute des Vorgangs

  SetCaseAttributes oHLCase, oMailRequest

  ' Gebe den Vorgang für alle User frei

  oCase.Unreserve

  ' save it to the helpline server

  oCase.Save

  ' Setze die Report Information

  Dim CaseRefNumber
  CaseRefNumber = oHLCase.GetValue("CASEINFO.REFERENCENUMBER", 0, 0, 0, 0)

  Dim sReportText
  sReportText = sReportText & vbLf & "CaseType:" & CStr(sCaseType)
  sReportText = sReportText & vbLf & "case-id:" & CStr(oHLCase.GetID)
  sReportText = sReportText & vbLf & "case-ref:" & CStr(CaseRefNumber)


  createCaseFromMail = sReportText
End Function

'---------------------------------------------------------------------------------------- extractRefNumber ---
Public Function ExtractRefNumber(ByRef subject)

  Dim refNum
  refNum = ""

  Dim startPos, endPos
  startPos = InStr(1, subject, "[#", 1)
  IF (startPos > 0) THEN
    startPos = startPos + 2
    ' skip "[#"

    endPos = InStr(startPos, subject, "]", 1)

    IF (endPos > 0) THEN
      refNum = Mid(subject, startPos, endPos - startPos)
    END IF
  END IF

  extractRefNumber = refNum
End Function


'---------------------------------------------------------------------------------------- extendCaseFromMail ---
Public Function ExtendCaseFromMail(ByRef oMailRequest, ByRef oCaseCfg, ByRef oHLServer, ByRef refNumber)

  LogText "extendCaseFromMail"

  Dim SearchCondition
  SearchCondition = "CASEINFO.REFERENCENUMBER= " & refNumber

  Dim cases
  Set cases = oHLServer.find_Cases(SearchCondition, 0)

  LogText "cases:" & cases.count

  Dim oCase
  For Each oCase In cases
    ExtendCase oCase, oMailRequest, oCaseCfg, oHLServer

    LogText "case extended"
    LogText "case-id:" & oCase.getHLObject.getID
    LogText "case-ref:" & CStr(refNumber)
  Next

  extendCaseFromMail = ""
End Function
'---------------------------------------------------------------------------------------- ExtendCase ---
Public Sub ExtendCase(ByRef ocase, ByRef oMailRequest, ByRef oCaseCfg, ByRef oHLServer)

  oCase.createSU

  AssociateSenderToCase oMailRequest, oCaseCfg, oHLServer, oCase

  SetSUAttributes ocase.getHLObject, oMailRequest

  oCase.mergeSUs

End Sub
'---------------------------------------------------------------------------------------- IsWorkflowEmail ---
Public Function IsWFEmail(ByRef subject, ByRef keywordList)
  LogText("IsWFEmail called")
  Dim item
  For Each item In keywordList
    IF (InStr(1, subject, item, 1) > 0) THEN
      LogText("IsWFEmail - " & item)
      IsWFEmail = True
      Exit For
    ELSE
      LogText("IsNotWFEmail - " & item)
      IsWFEmail = False
    END IF
  Next
End Function
'---------------------------------------------------------------------------------------- main ---
ProcessIn
' call the main ü entry point
