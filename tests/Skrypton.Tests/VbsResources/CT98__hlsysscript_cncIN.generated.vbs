Option Explicit

'--------------------------------------------------------------------------------------- ProcessIn ---
Public Sub ProcessIn()
  LogText "ProcessRequestMail start."

  Dim mailRequest

  Set mailRequest = session("mailrequest")

  LogText "mail subject: " & mailRequest.subject
  LogText "mail To: " & mailRequest.To

  IF IsAutoReplyMail(mailRequest.Subject) THEN
    LogText "Out of Office AutoReply"
    Exit Sub
  END IF

  Dim extendCaseSuccess
  extendCaseSuccess = TryExtendCase(mailRequest.Subject)
  IF extendCaseSuccess = False THEN
    LogText "Extend case failed. Start new process"
    IF IsFMMail(mailRequest.To) THEN
      StartNewFMWorkflow(mailRequest.Subject)
    ELSEIF IsHRMail(mailRequest.To) THEN
      StartNewHRWorkflow(mailRequest.Subject)
    ELSE
      StartNewWorkflow(mailRequest.Subject)
    END IF
  END IF

  LogText "ProcessRequestMail end."
End Sub

'--------------------------------------------------------------------------------------- IsAutoReplyMail ---
Public Function IsAutoReplyMail(ByRef mailSubject)
  Dim autoReplyList
  Dim item
  Dim retVal
  retVal = False
  autoReplyList = Array("Out of Office:", "Abwesend:")

  For Each item In autoReplyList
    IF (InStr(1, mailSubject, item, 1) > 0) THEN
      retVal = True
    END IF
  Next
  IsAutoReplyMail = retVal
End Function

'--------------------------------------------------------------------------------------- TryExtendCase ---
Public Function TryExtendCase(ByRef mailSubject)
  Dim refNumber
  Dim caseToExtend
  Dim reportText
  Dim retVal
  retVal = False

  refNumber = ExtractRefNumber(mailSubject)
  IF (Len(refNumber) > 0) THEN
    LogText "RefNumber > 0"
    Set caseToExtend = session.GetCaseByReferenceNumber(refNumber)
    IF (session.CanExtendWorkflowCase(caseToExtend)) THEN
      LogText "CanExtend"
      reportText = session.DoExtendWorkflowCase(caseToExtend)
      LogText reportText
      retVal = True
    END IF
  END IF
  TryExtendCase = retVal
End Function

'--------------------------------------------------------------------------------------- StartNewWorkflow ---
Public Sub StartNewWorkflow(ByRef mailSubject)
  Dim imKeywords
  Dim rfKeywords
  Dim cmKeywords
  Dim fmKeywords
  Dim hrKeywords
  Dim reportText
  rfKeywords = Array("[ServiceRequest]", "Anfrage", "request", "Frage", "question")
  imKeywords = Array("[Incident]", "Incident", "Störung", "Hilfe", "help")
  cmKeywords = Array("[RFC]", "Änderung", "Change")
  fmKeywords = Array("[Facility]", "Haustechnik", "FM")
  hrKeywords = Array("[HR]", "Personal")

  IF (IsWFEmail(mailSubject, rfKeywords) = True) THEN
    reportText = session.NewWorkflowFromMail("RequestFulfillment")
    LogText reportText
    Exit Sub
  END IF
  IF (IsWFEmail(mailSubject, imKeywords) = True) THEN
    reportText = session.NewWorkflowFromMail("IncidentManagement")
    LogText reportText
    Exit Sub
  END IF
  IF (IsWFEmail(mailSubject, cmKeywords) = True) THEN
    reportText = session.NewWorkflowFromMail("ChangeManagement")
    LogText reportText
    Exit Sub
  END IF
  IF (IsWFEmail(mailSubject, fmKeywords) = True) THEN
    reportText = session.NewWorkflowFromMail("FacilityIncidentManagement")
    LogText reportText
    Exit Sub
  END IF
  IF (IsWFEmail(mailSubject, hrKeywords) = True) THEN
    reportText = session.NewWorkflowFromMail("HRRequestManagement")
    LogText reportText
    Exit Sub
  END IF
  reportText = session.NewWorkflowFromMail("Request")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- StartNewFMWorkflow ---
Public Sub StartNewFMWorkflow(ByRef mailSubject)
  Dim reportText

  reportText = session.NewWorkflowFromMail("FacilityIncidentManagement")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- StartNewHRWorkflow ---
Public Sub StartNewHRWorkflow(ByRef mailSubject)
  Dim reportText

  reportText = session.NewWorkflowFromMail("HRRequestManagement")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- LogText ---
Public Sub LogText(ByRef sText)
  'Uncomment to enable logging
  session("processtext") = session("processtext") & sText & vbNewLine
End Sub

'---------------------------------------------------------------------------------------- ExtractRefNumber ---
Public Function ExtractRefNumber(ByRef mailSubject)
  Dim refNum
  refNum = ""
  Dim startPos
  Dim endPos

  startPos = InStr(1, mailSubject, "[#", 1)
  IF (startPos > 0) THEN
    startPos = startPos + 2
    ' skip "[#"
    endPos = InStr(startPos, mailSubject, "]", 1)
    IF (endPos > 0) THEN
      refNum = Mid(mailSubject, startPos, endPos - startPos)
    END IF
  END IF
  extractRefNumber = refNum
End Function

'--------------------------------------------------------------------------------------- IsFMMail ---
Public Function IsFMMail(ByRef mailTo)
  LogText("IsFMMail called")
  Dim retVal
  retVal = False
  IF mailTo = "haustechnik@helplinedemo.de" THEN
    retVal = True
  END IF

  IsFMMail = retVal
End Function

'--------------------------------------------------------------------------------------- IsFMMail ---
Public Function IsHRMail(ByRef mailTo)
  LogText("IsHRMail called")
  Dim retVal
  retVal = False
  IF mailTo = "personal@helplinedemo.de" THEN
    retVal = True
  END IF

  IsHRMail = retVal
End Function

'---------------------------------------------------------------------------------------- IsWorkflowEmail ---
Public Function IsWFEmail(ByRef mailSubject, ByRef keywordList)
  LogText("IsWFEmail called")
  Dim item
  Dim retVal
  retVal = False

  For Each item In keywordList
    IF (InStr(1, mailSubject, item, 1) > 0) THEN
      LogText("IsWFEmail - " & item)
      retVal = True
      Exit For
    END IF
  Next
  IsWFEmail = retVal
End Function
'---------------------------------------------------------------------------------------- main ---
ProcessIn
