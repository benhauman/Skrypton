Option Explicit

'--------------------------------------------------------------------------------------- ProcessIn ---
Sub ProcessIn
  LogText "ProcessRequestMail start."

  Dim mailRequest
     
  Set mailRequest = session("mailrequest")
    
  LogText "mail subject: " & mailRequest.subject
  LogText "mail To: " & mailRequest.To
    
  If IsAutoReplyMail(mailRequest.Subject) Then
    LogText "Out of Office AutoReply"     
    Exit Sub
  End If
  
  Dim extendCaseSuccess
  extendCaseSuccess = TryExtendCase(mailRequest.Subject)
  If extendCaseSuccess = False Then
    LogText "Extend case failed. Start new process"
    If IsFMMail(mailRequest.To) Then
      StartNewFMWorkflow(mailRequest.Subject)
    ElseIf IsHRMail(mailRequest.To) Then
      StartNewHRWorkflow(mailRequest.Subject)
    Else  
      StartNewWorkflow(mailRequest.Subject)
    End If    
  End If
    
  LogText "ProcessRequestMail end."
End Sub

'--------------------------------------------------------------------------------------- IsAutoReplyMail ---
Function IsAutoReplyMail(mailSubject)
  Dim autoReplyList
  Dim item
  Dim retVal : retVal = False
  autoReplyList = Array("Out of Office:", "Abwesend:")
  
  For Each item in autoReplyList
    If (InStr(1, mailSubject, item, 1) > 0) Then 
      retVal = True
    End If
  Next
  IsAutoReplyMail = retVal
End Function

'--------------------------------------------------------------------------------------- TryExtendCase ---
Function TryExtendCase(mailSubject)
  Dim refNumber
  Dim caseToExtend
  Dim reportText
  Dim retVal : retVal = False
  
  refNumber = ExtractRefNumber(mailSubject)
  If (Len(refNumber) > 0) Then
    LogText "RefNumber > 0"
    Set caseToExtend = session.GetCaseByReferenceNumber(refNumber)
    If (session.CanExtendWorkflowCase(caseToExtend)) Then
      LogText "CanExtend"
      reportText = session.DoExtendWorkflowCase(caseToExtend)
      LogText reportText
      retVal = True
    End If 
  End If
  TryExtendCase = retVal
End Function

'--------------------------------------------------------------------------------------- StartNewWorkflow ---
Sub StartNewWorkflow(mailSubject)
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
  
  If (IsWFEmail(mailSubject, rfKeywords) = True) Then
    reportText = session.NewWorkflowFromMail("RequestFulfillment")
    LogText reportText
    Exit Sub  
  End If
  If (IsWFEmail(mailSubject, imKeywords) = True) Then
    reportText = session.NewWorkflowFromMail("IncidentManagement")
    LogText reportText
    Exit Sub
  End If
  If (IsWFEmail(mailSubject, cmKeywords) = True) Then
    reportText = session.NewWorkflowFromMail("ChangeManagement")
    LogText reportText
    Exit Sub
  End If
  If (IsWFEmail(mailSubject, fmKeywords) = True) Then
    reportText = session.NewWorkflowFromMail("FacilityIncidentManagement")
    LogText reportText
    Exit Sub
  End If
  If (IsWFEmail(mailSubject, hrKeywords) = True) Then
    reportText = session.NewWorkflowFromMail("HRRequestManagement")
    LogText reportText
    Exit Sub
  End If  
  reportText = session.NewWorkflowFromMail("Request")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- StartNewFMWorkflow ---
Sub StartNewFMWorkflow(mailSubject)
  Dim reportText

  reportText = session.NewWorkflowFromMail("FacilityIncidentManagement")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- StartNewHRWorkflow ---
Sub StartNewHRWorkflow(mailSubject)
  Dim reportText

  reportText = session.NewWorkflowFromMail("HRRequestManagement")
  LogText reportText
End Sub

'--------------------------------------------------------------------------------------- LogText ---
Sub LogText(sText)
  'Uncomment to enable logging
  session("processtext") = session("processtext") & sText & vbNewLine
End Sub

'---------------------------------------------------------------------------------------- ExtractRefNumber ---
Function ExtractRefNumber(mailSubject)
  Dim refNum : refNum = ""
  Dim startPos 
  Dim endPos
  
  startPos = InStr(1, mailSubject, "[#", 1)
  If ( startPos > 0 ) Then
    startPos = startPos + 2 ' skip "[#" 
    endPos = InStr( startPos, mailSubject, "]", 1)
    If ( endPos > 0) Then
      refNum = Mid(mailSubject, startPos, endPos - startPos)
    End If
  End If
  extractRefNumber = refNum
End Function

'--------------------------------------------------------------------------------------- IsFMMail ---
Function IsFMMail(mailTo)
  LogText("IsFMMail called")
  Dim retVal : retVal = False
  If mailTo = "haustechnik@helplinedemo.de" Then
    retVal = True
  End If

  IsFMMail = retVal
End Function

'--------------------------------------------------------------------------------------- IsFMMail ---
Function IsHRMail(mailTo)
  LogText("IsHRMail called")
  Dim retVal : retVal = False
  If mailTo = "personal@helplinedemo.de" Then
    retVal = True
  End If

  IsHRMail = retVal
End Function

'---------------------------------------------------------------------------------------- IsWorkflowEmail ---
Function IsWFEmail(mailSubject, keywordList)
    LogText("IsWFEmail called")
    Dim item
    Dim retVal : retVal = False
  
    For Each item in keywordList
      If (InStr(1, mailSubject, item, 1) > 0) Then 
        LogText("IsWFEmail - " & item)
        retVal = True
        Exit For
      End If
    Next
  IsWFEmail = retVal
End Function
'---------------------------------------------------------------------------------------- main ---
ProcessIn
