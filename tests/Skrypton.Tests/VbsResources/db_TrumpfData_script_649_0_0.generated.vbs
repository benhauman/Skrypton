'----------------------------------------------------------------------------------------------------------
'Globale Funktion zur Initialisierung der Datei hlStartITIL2.dll.
'Diese DLL-Datei beinhaltet alle globalen Funktionen und Prozeduren,
'die innerhalb der Start ITIL Konfiguration verwendet werden.
'Diese Funktion darf nicht aus dem gloabeln Script entfernt werden !
'Global Function for initializing the file hlStartITIL2.dll.
'This dll file contains any global functions and subs used for the
'Start ITIL configuration.
'Do not remove this function from the global script !
'Copyright (C) 1994-2006 PMCS GmbH & Co.
Public Function hlITIL2()
  Set hlITIL2 = CreateObject("hlStartITIL2.Global")
  hlITIL2.SelfCheck hlContext
End Function
'----------------------------------------------------------------------------------------------------------
'Globale Konstanten fuer freie Assoziationsdefinitionen
Const Skrypton.LegacyParser.Tokens.Basic.NameToken:HLASC_SoftwareLicenseFolderView = Skrypton.LegacyParser.Tokens.Basic.NumericValueToken:110941

Const Skrypton.LegacyParser.Tokens.Basic.NameToken:HLASC_SoftwareLicenseGroupView = Skrypton.LegacyParser.Tokens.Basic.NumericValueToken:110944

'----------------------------------------------------------------------------------------------------------
'Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
Public Sub Trace(ByRef hlContext, ByRef text)
  hlContext.trace 1, text
End Sub
'----------------------------------------------------------------------------------------------------------
'Funktion InfoMail
'Zum Aufrufen aus EBL-Skripten von Vorgaengen
Public Sub InfoMail(ByRef hlContext, ByRef hlCase, ByRef Subject, ByRef MailSender, ByRef Receiver, ByRef CC, ByRef body, ByRef SendAttachments)

  Dim Email
  Set Email = hlContext.CreateMail

  'Falls der Parameter <SendAttachmnets> beim Aufruf "1" ist, werden Anhaenge mitversandt
  IF CBool(SendAttachments) = True THEN
    Dim AttachIDs
    Dim AttachID
    Dim Attachment
    Set Attachment = Nothing
    AttachIDs = hlCase.GetAttachmentKeys("HLOBJECTINFO.ATTACHMENT", 0)
    For Each AttachID In AttachIDs
      Set Attachment = hlCase.GetAttachment("HLOBJECTINFO.ATTACHMENT", AttachID, 0)
      IF Attachment.Size > 0 THEN
        Dim MailAttachment
        Set MailAttachment = Nothing
        Set MailAttachment = Email.AddAttachment
        MailAttachment.name = Attachment.name
        MailAttachment.data = Attachment.data
      END IF
    Next
  END IF

  IF MailSender < > "" THEN
    Email.SenderMail = MailSender
  END IF
  Email.To = Receiver
  Email.Subject = Subject
  Email.Body = body
  IF CC < > "" THEN
    Email.CC = CC
  END IF
  Call hlContext.SendRequestMail(Email)
End Sub

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
Public Sub CreateSubject(ByRef hlContext, ByRef Survey, ByRef hlCaller)
  Dim language
  language = hlCaller.GetValue("PersonGeneral.Language", 0, 0, 0, 0)
  IF language = "LanguageGerman" THEN
    Call Survey.SetValue("CaseGeneral.Subject", 0, 0, 0, "Umfrage zur Service-Leistung ihres Support-Teams")
  ELSE
    Call Survey.SetValue("CaseGeneral.Subject", 0, 0, 0, "Survey about the Service-Quality from your Support-Team")
  END IF
End Sub
'----------------------------------------------------------------------------------------------------------
Public Sub InviteSurveyEmail(ByRef hlContext, ByRef hlCase, ByRef hlCaller)
  'Email an den Anfrager eines Survey-Vorgangs, um diesen zur Teilnahme an der
  'Umfrage aufzufordern.
  'Email to Requester of a Survey-Case to invite him to take part on the survey
  Dim SUIDx
  SUIDx = hlITIL2.GetLastSUIdx(hlCase, hlContext)
  Dim MailRequest
  MailRequest = hlCase.GetValue("CaseGeneral.DefaultNotification", 0, 0, 0, 0)
  IF MailRequest = "DefaultNotificationEmail" And SUIDx = 1 THEN
    Dim strCRLF
    strCRLF = CHR(13) & CHR(10)
    Dim Creationdate, Datum, subject, body
    Dim refnumber
    refnumber = hlcase.GetValue("CASEINFO.REFERENCENUMBER", 0, 0, 0, 0)
    Dim portallink
    portallink = "http://localhost/helplineportal/"
    Dim surname
    surname = hlCaller.GetValue("PersonGeneral.PersonSurname", 0, 0, 0, 0)
    Dim letteraddress
    letteraddress = hlCaller.GetValue("PersonGeneral.ShortLetterAddress", 0, 0, 0, 0)
    Dim Anrede
    Anrede = "Sehr geehrte Damen und Herren,"
    Dim PersonAddress
    PersonAddress = "Dear Mrs./Ms. or Mr.,"
    Dim language
    language = hlCaller.GetValue("PersonGeneral.Language", 0, 0, 0, 0)

    IF language = "LanguageGerman" THEN
      IF letteraddress = "" THEN
        letteraddress = "Herr/Frau"
      END IF
      Anrede = "Sehr geehrte(r) " & CStr(letteraddress) & " " & CStr(surname) & ","

      'Hier wird die Betreffzeile erstellt
      'The subject field is entered here
      Creationdate = hlcase.GetValue("HLOBJECTINFO.CREATIONTIME", 7, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      subject = "Umfrage zur Service-Leistung ihres Support-Teams"

      'Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
      'If the case was recorded, the requester receives the according information
      body = CStr(Anrede)
      body = body & strCRLF & strCRLF
      body = body & "Wir wollen besser werden!"
      body = body & strCRLF & "Dazu benoetigen wir Ihre Unterstuetzung und Ihr Feedback."
      body = body & strCRLF & strCRLF
      body = body & "Sie wurden am " & Datum & " durch ein Zufallsverfahren ausgewaehlt, an einer Umfrage zu den Service-Leistungen Ihres Support-Teams teilzunehmen."
      body = body & strCRLF & strCRLF
      body = body & "Die Teilnahme ist freiwillig und erfolgt ueber das helpLine Portal."
      body = body & strCRLF & strCRLF
      body = body & "Rufen Sie im Browser bitte folgende URL auf:"
      body = body & strCRLF & portallink & strCRLF & strCRLF
      body = body & "Klicken Sie unter 'Ihre Anfragen' auf den Eintrag 'Umfragen'. "
      body = body & "Dort finden Sie das Umfrage-Formular mit der Nummer " & refnumber & ". "
      body = body & strCRLF & strCRLF
      body = body & "Wir freuen uns sehr, wenn Sie sich die Zeit nehmen, die Fragen zu beantworten."
      body = body & strCRLF & strCRLF
      body = body & "Wir bedanken uns fuer Ihre Unterstuetzung!"
      body = body & strCRLF & strCRLF
      body = body & strCRLF & "Mit freundlichen Gruessen"
      body = body & strCRLF & strCRLF
      body = body & "Ihr Support Team"
    ELSE
      IF letteraddress = "" THEN
        letteraddress = "Mrs./Ms./Mr."
      END IF
      PersonAddress = "Dear " & CStr(letteraddress) + " " & CStr(surname) & ","

      'Hier wird die Betreffzeile erstellt
      'The subject field is entered here
      Creationdate = hlcase.GetValue("HLOBJECTINFO.CREATIONTIME", 7, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      subject = "Survey about the Service-Quality from your Support-Team"


      'Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
      'If the case was recorded, the requester receives the according information
      body = CStr(PersonAddress)
      body = body & strCRLF & strCRLF
      body = body & "We would like to improve the efficiency of Service-Support!"
      body = body & strCRLF & "Therefore we need your assistance and your feedback."
      body = body & strCRLF & strCRLF
      body = body & "You where chosen by random on " & Datum & " to take part on the Survey."
      body = body & strCRLF & strCRLF
      body = body & "The participation on the survey is voluntarily. You can take part on the survey via the helpLine Portal."
      body = body & strCRLF & strCRLF
      body = body & "Start your Browser and choose the following URL:"
      body = body & strCRLF & portallink & strCRLF & strCRLF
      body = body & "Then klick 'Survey' in the menue 'Your Requests'. " & strCRLF
      body = body & "There, you will find the Questionnaire with the reference number " & refnumber & ". "
      body = body & strCRLF & strCRLF
      body = body & "It would be nice, if you invest your time to response the questions."
      body = body & strCRLF & strCRLF
      body = body & "We thank you for your assistance!"
      body = body & strCRLF & strCRLF
      body = body & strCRLF & "With best regards"
      body = body & strCRLF & strCRLF
      body = body & "Yours Support Team"
    END IF

    Dim Email
    Set Email = hlContext.CreateMail

    'Ermittle die Emailadresse des Anfragers
    'Detect email adress of requester
    Dim Emailadress
    Emailadress = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
    IF Emailadress = "" THEN
      Emailadress = "Username@yourcompany.com"
      subject = "Diese EMail konnte nicht zugestellt werden"
      body = "Die Mail fuer die Anfragenummer "
      body = body & hlcase.GetValue("CASEINFO.REFERENCENUMBER", 0, 0, 0, 0)
      body = body & " konnte wegen einer fehlenden E-Mail Adresse nicht zugestellt werden."
    END IF
    Email.To = Emailadress
    Email.Subject = subject
    Email.Body = body
    Call hlContext.SendRequestMail(Email)
  END IF
End Sub
'----------------------------------------------------------------------------------------------------------
'Diese Funktion steuert den SystemTask wenn dieser im Vorgangstyp Task konfiguriert wurde.
'This function controls a SystemTask if it had been configured within the casetype Task.
Public Sub MyTask1(ByRef hlContext)
  Dim hlObj
  Set hlObj = hlContext.GetCurrentObject()
  Dim lcid
  lcid = 0
  lcid = hlContext.GetLocaleID
  Dim LangID
  LangID = 0
  LangID = hlContext.LangIDFromLCID(lcid)

  'Gesetzte Daten aus dem aktuellen Task auslesen, diese werden dem zu erzeugenden Systemtask mitgegeben.
  'Read setted data of current task and take them into the created Systemtask.
  Dim Priority
  Priority = hlObj.GetValue("CaseClassificationAttribute.Priority", 0, 0, 0, 0)
  Dim TaskType
  TaskType = hlObj.GetValue("TaskGeneral.TaskType", 0, 0, 0, 0)
  Dim Subject
  Subject = hlObj.GetValue("TaskGeneral.Subject", 0, 0, 0, 0)
  Dim Description
  Description = hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0)
  Dim ExOperation
  ExOperation = hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, 0, 0)
  Dim AssignedGroup
  AssignedGroup = hlObj.GetValue("CaseSpecialRouting.AssignedGroup", 0, 0, 0, 0)
  Dim AssignedPerson
  AssignedPerson = hlObj.GetValue("CaseSpecialRouting.AssignedPerson", 0, 0, 0, 0)
  Dim Team
  Team = hlObj.GetValue("Keywords.KeywordOrga", 0, 0, 0, 0)
  Dim newTask
  Set newTask = hlContext.createobject("Task")

  newTask.SetValue "CaseClassificationAttribute.Priority", 0, 0, 0, Priority
  newTask.SetValue "TaskGeneral.TaskType", 0, 0, 0, TaskType
  newTask.SetValue "TaskGeneral.Subject", 0, 0, 0, Subject
  newTask.SetValue "Keywords.KeywordOrga", 0, 0, 0, Team

  Dim hasContent
  hasContent = hlObj.HasContent("TaskDesignWorkflow.TaskWorkflow_CA", 0, 0)
  IF hasContent < > 0 THEN
    Dim contentIDs
    Dim contentID
    Dim newContentID
    Dim assignedGroupWF
    Dim assignedPersonWF
    Dim descriptionWF
    Dim subjectWF
    subjectWF = hlObj.GetValue("TaskDesignWorkflow.FlagWorkflowSubject", 0, 0, 0, 0)
    newTask.SetValue "TaskDesignWorkflow.FlagWorkflowSubject", 0, 0, 0, subjectWF
    contentIDs = hlObj.GetContentIDs("TaskDesignWorkflow.TaskWorkflow_CA", 0)
    newTask.SetValue "TaskWorkflowAttribute.WorkflowStep", 0, 0, 0, 1
    For Each contentID In contentIDs
      assignedGroupWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup", 0, contentID, 0, 0)
      assignedPersonWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson", 0, contentID, 0, 0)
      descriptionWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText", 0, contentID, 0, 0)
      newContentID = hlObj.GenerateContentID
      newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup", 0, newContentID, 0, assignedGroupWF
      newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson", 0, newContentID, 0, assignedPersonWF
      newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText", 0, newContentID, 0, descriptionWF
    Next
  END IF


  Dim Assets
  Dim Asset
  Assets = hlObj.GetItemsEx(0, 0, 131)
  For Each Asset In Assets
    Call newTask.AddItemEx(0, Asset, 0, 131)
  Next
  Dim RefNumber
  RefNumber = hlObj.GetValue("CASEINFO.REFERENCENUMBER", 0, 0, 0, 0)
  IF LangID = 7 THEN
    Description = Description & vbNewLine & vbNewLine & "[Diese Aufgabe wurde automatisch durch den Systemtask mit der Bezugsnummer '" & RefNumber & " erstellt.]"
  ELSE
    Description = Description & vbNewLine & vbNewLine & "[This Task was created automatically by Systemtask with the Reference Number '" & RefNumber & "'.]"
  END IF
  newTask.SetValue "CaseDescription.DescriptionText", 0, 0, 0, Description
  newTask.SetValue "CaseDiagnosis.DiagnosisText", 0, 0, 0, ExOperation
  newTask.SetValue "CaseSpecialRouting.AssignedGroup", 0, 0, 0, AssignedGroup
  newTask.SetValue "CaseSpecialRouting.AssignedPerson", 0, 0, 0, AssignedPerson
  newTask.SetValue "Keywords.KeywordOrga", 0, 0, 0, Team
  hlContext.SaveObject(newTask)
  Call newTask.Unreserve()
End Sub

'Festlegung der Definitionen eines SystemTasks pro Tag.
'Determining of definitions of a SystemTask by day.
Public Sub CreateSystemTaskDefbyDay(ByRef SysTaskBeginnDate, ByRef SysTaskEndDate, ByRef NoEndDate, ByRef NumberOfDays, ByRef taskDefname, ByRef recurrenceEndType)
  Dim hlObj
  Set hlObj = hlContext.GetCurrentObject()
  Dim hlSystemTask
  Set hlSystemTask = hlContext.CreateSystemTask(0)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Dim systemTaskDefinitionName
  systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, 0)
  Dim scriptCode
  scriptCode = "MyTask1"
  Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE", 0, 0, 0, SysTaskBeginnDate)
  'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
  'Check which option had been chosen in duration of the SystemTask.
  Dim newTaskEndTime
  '=No EndDate
  'Alt - Anfang
  'If recurrenceEndType = "0" Then
  '	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
  'Else
  '=UserEndDate
  '	If recurrenceEndType = "2"	Then
  '		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
  '	End If
  'End If
  'Alt - Ende
  'Neu - Anfang
  IF recurrenceEndType = "0" THEN
    Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskBeginnDate)
    recurrenceEndType = "1"
  ELSE
    IF recurrenceEndType = "2" THEN
      Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskEndDate)
      recurrenceEndType = "1"
    END IF
  END IF
  'Neu - Ende

  Call hlSystemTask.SetValue("SYSTASKINFO.ENDTYPE", 0, 0, 0, recurrenceEndType)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE", 0, 0, 0, scriptCode)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL", 0, 0, 0, NumberOfDays)
  Call hlContext.SaveSystemTask(hlSystemTask)
  Dim hlSystemTaskDefinitionObj
  Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
  Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub

'Entfernt einen vorhandenen SystemTask.
'Remove an existing SystemTask.
Public Sub DeleteSystemTask(ByRef hlContext, ByRef hlObj, ByRef hlSystemTask, ByRef taskname)
  Call hlContext.RemoveSystemTask(hlSystemTask)
End Sub

'Festlegung der Definitionen eines SystemTasks pro Woche.
'Determining of definitions of a SystemTask by week.
Public Sub CreateSystemTaskDefbyWeek(ByRef SysTaskBeginnDate, ByRef SysTaskEndDate, ByRef NoEndDate, ByRef NumberOfWeeks, ByRef MondayFlag, ByRef TuesdayFlag, ByRef WednesdayFlag, ByRef ThursdayFlag, ByRef FridayFlag, ByRef SaturdayFlag, ByRef SundayFlag, ByRef taskDefname, ByRef recurrencedaymask, ByRef recurrenceEndType)
  Dim hlObj
  Set hlObj = hlContext.GetCurrentObject()
  Dim hlSystemTask
  Set hlSystemTask = hlContext.CreateSystemTask(0)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Dim systemTaskDefinitionName
  systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, 0)
  Dim scriptCode
  scriptCode = "MyTask1"

  'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
  'Check which option had been chosen in duration of the SystemTask.
  'Alt - Anfang
  'If recurrenceEndType = "0" Then
  '	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
  'Else
  '	If recurrenceEndType = "2"	Then
  '		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
  '		recurrenceEndType = "1"
  '	End If
  'End If
  'Alt - Ende
  'Neu - Anfang
  IF recurrenceEndType = "0" THEN
    Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskBeginnDate)
    recurrenceEndType = "1"
  ELSE
    IF recurrenceEndType = "2" THEN
      Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskEndDate)
      recurrenceEndType = "1"
    END IF
  END IF
  'Neu - Ende
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE", 0, 0, 0, scriptCode)
  Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE", 0, 0, 0, SysTaskBeginnDate)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE", 0, 0, 0, recurrenceEndType)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL", 0, 0, 0, NumberOfWeeks)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.DAYMASK", 0, 0, 0, recurrencedaymask)

  Call hlContext.SaveSystemTask(hlSystemTask)
  Dim hlSystemTaskDefinitionObj
  Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
  Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub

'Festlegung der Definitionen eines SystemTasks pro Monat.
'Determining of definitions of a SystemTask by month.
Public Sub CreateSystemTaskDefbyMonth(ByRef SysTaskBeginnDate, ByRef SysTaskEndDate, ByRef NoEndDate, ByRef DayOfMonth, ByRef NumberOfMonths, ByRef taskDefname, ByRef recurrenceEndType)
  Dim hlObj
  Set hlObj = hlContext.GetCurrentObject()
  Dim hlSystemTask
  Set hlSystemTask = hlContext.CreateSystemTask(0)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Dim systemTaskDefinitionName
  systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, 0)
  Dim scriptCode
  scriptCode = "MyTask1"

  'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
  'Check which option had been chosen in duration of the SystemTask.
  'Alt - Anfang
  'If recurrenceEndType = "0" Then
  '	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
  'Else
  '	If recurrenceEndType = "2"	Then
  '		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
  '	End If
  'End If
  'Alt - Ende
  'Neu - Anfang
  IF recurrenceEndType = "0" THEN
    Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskBeginnDate)
    recurrenceEndType = "1"
  ELSE
    IF recurrenceEndType = "2" THEN
      Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, SysTaskEndDate)
      recurrenceEndType = "1"
    END IF
  END IF
  'Neu - Ende
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE", 0, 0, 0, recurrenceEndType)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.DAYOFMONTH", 0, 0, 0, DayOfMonth)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INSTANCE", 0, 0, 0, "0")
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL", 0, 0, 0, NumberOfMonths)
  Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE", 0, 0, 0, SysTaskBeginnDate)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE", 0, 0, 0, scriptCode)
  Call hlContext.SaveSystemTask(hlSystemTask)
  Dim hlSystemTaskDefinitionObj
  Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
  Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub

'Sub fuehrt den SystemTask einmalig aus.
'Sub execute SystemTask one-time.
Public Sub CreateOneTimeSystemTask(ByRef OneTimeTask, ByRef SysTaskEndDate, ByRef SysTaskBeginnDate, ByRef taskDefname)
  Dim hlObj
  Set hlObj = hlContext.GetCurrentObject()
  Dim hlSystemTask
  Set hlSystemTask = hlContext.CreateSystemTask("0")
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Dim scriptCode
  scriptCode = "MyTask1"

  'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
  'Check which option had been chosen in duration of the SystemTask.
  Dim systemTaskDefinitionName
  systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, 0)
  Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME", 0, 0, 0, taskDefname)
  Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE", 0, 0, 0, SysTaskBeginnDate)
  Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE", 0, 0, 0, "09.09.2099 09:09:09")
  Call hlSystemTask.SetValue("SYSTASKINFO.ENDTYPE", 0, 0, 0, 1)
  Call hlSystemTask.SetValue("SYSTASKINFO.ENDCOUNT", 0, 0, 0, 1)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE", 0, 0, 0, 0)
  Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL", 0, 0, 0, 1)
  Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE", 0, 0, 0, scriptCode)

  'Wenn kein Datum angegebene wurde, muss eine Fehlermeldung angezeigt werden.
  'If no date was entered, show an error message.
  IF SysTaskBeginnDate = "" THEN
    errCode = "#ERR_TSKMNT_002"
  END IF
  Call hlContext.SaveSystemTask(hlSystemTask)
  Dim hlSystemTaskDefinitionObj
  Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
  Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
  Call hlObj.HasSystemTask(systemTaskDefinitionName)
End Sub
'----------------------------------------------------------------------------------------------------------
Public Function IsValidObject(ByRef obj)
  IsValidObject =(IsObject(obj) And(Not(obj Is Nothing)))
End Function

'XML-Export Neuanlage

Public Sub ExportObject(ByRef hlContext, ByRef hlObj)
  Dim objDefname
  objDefname = hlObj.GetType()
  Dim aliasname
  aliasname = "NewCI" & objDefname
  Dim NewChangeObj
  NewChangeObj = hlObj.GetValue("TrumpfAssetGeneral.DataToSAPAMChange", 0, 0, 0, 0)
  IF NewChangeObj = "0" Or NewChangeObj = "" THEN
    aliasname = aliasname
  ELSE
    aliasname = "ChangedCI" & objDefname
  END IF

  ' VBScript source code
  Dim xmldoc
  Set xmldoc = CreateObject("msxml2.DomDocument")

  'create root element
  Dim nodeData
  Set nodeData = xmldoc.appendChild(xmldoc.createElement("Data"))
  Dim nodeObjects
  Set nodeObjects = nodeData.appendChild(xmldoc.createElement("Objects"))
  Dim nodeObject
  Set nodeObject = nodeObjects.appendChild(xmldoc.createElement(objDefname))
  Dim attAliasName
  Set attAliasName = xmldoc.createAttribute("aliasname")
  attAliasName.Text = aliasname
  nodeObject.Attributes.setNamedItem attAliasName
  Dim nodeAttributes
  Set nodeAttributes = nodeObject.appendChild(xmldoc.createElement("Attributes"))
  Dim nodeRelations
  Set nodeRelations = nodeData.appendChild(xmldoc.createElement("Relations"))
  '///////////////////////////////////////////////////////////

  '//////////////// HLOBJECT.ID
  'Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "AssetGeneral.AssetName", hlObj.GetValue("AssetGeneral.AssetName", 0, 0, 0, 0))
  ' hlObj.GetValue("AssetGeneral.AssetName", 0,0,0,0)
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "AccountingDetail.CostCenter", hlObj.GetValue("AccountingDetail.CostCenter", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "AssetGeneral.Serialnumber", hlObj.GetValue("AssetGeneral.Serialnumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "ProcurementDetail.AllocationNumber", hlObj.GetValue("ProcurementDetail.AllocationNumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "ProcurementDetail.AllocationType", hlObj.GetValue("ProcurementDetail.AllocationType", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "ProcurementDetail.OrderNumber", hlObj.GetValue("ProcurementDetail.OrderNumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "ProcurementDetail.OrderPosition", hlObj.GetValue("ProcurementDetail.OrderPosition", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "ProcurementDetail.VendorNumber", hlObj.GetValue("ProcurementDetail.VendorNumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, True, "TrumpfAssetGeneral.CINumber", hlObj.GetValue("TrumpfAssetGeneral.CINumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.CompanyCode", hlObj.GetValue("TrumpfAssetGeneral.CompanyCode", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.InvestmentNumber", hlObj.GetValue("TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.Manufacturer", hlObj.GetValue("TrumpfAssetGeneral.Manufacturer", 0, 0, 0, 0))
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.SAPCostCenter", hlObj.GetValue("TrumpfAssetGeneral.SAPCostCenter", 0, 0, 0, 0))

  ' Save to File
  Dim Filename
  IF NewChangeObj = "0" Or NewChangeObj = "" THEN
    Filename = "\\srvditz1\pi_intern\helpline\HELPLINE_out\c11\hlnew" & objDefname & "_" & hlObj.GetID & ".xml"
  ELSE
    Filename = "\\srvditz1\pi_intern\helpline\HELPLINE_out\c11\hlchange" & objDefname & "_" & hlObj.GetID & ".xml"
  END IF
  xmldoc.Save(Filename)

End Sub

Public Sub AppendNode(ByRef hlContext, ByRef xmldoc, ByRef nodeObject, ByRef iskey, ByRef key, ByRef value)
  Dim valueNode
  Set valueNode = xmldoc.createElement(key)
  Dim cdata
  Set cdata = xmldoc.createCDATASection(value)
  valueNode.appendChild(cdata)
  nodeObject.appendChild(valueNode)

  Dim attIsKey
  Set attIsKey = xmldoc.createAttribute("iskey")
  IF (iskey) THEN
    attIsKey.Text = "true"
  ELSE
    attIsKey.Text = "false"
  END IF
  valueNode.Attributes.setNamedItem attIsKey

End Sub

'XML-Export Incident wegen Eleminierung

Public Sub ExportObjectIncident(ByRef hlContext, ByRef hlObj)
  Dim objDefname
  objDefname = "IncidentRequest"
  Dim aliasname1
  aliasname1 = "obj1"
  Dim aliasname2
  aliasname2 = "obj2"
  Dim aliasnameSU
  aliasnameSU = "objSU"
  Dim ElimierungsgrundDE
  ElimierungsgrundDE = hlObj.GetValue("TrumpfAssetStatus.CISubStatus", 7, 0, 0, 0)
  Dim ElimierungsgrundEN
  ElimierungsgrundEN = hlObj.GetValue("TrumpfAssetStatus.CISubStatus", 9, 0, 0, 0)
  Dim Buchungskreis
  Buchungskreis = hlObj.GetValue("TrumpfAssetGeneral.CompanyCode", 0, 0, 0, 0)
  Dim Buchungskreis1
  Buchungskreis1 = hlObj.GetValue("TrumpfAssetGeneral.CompanyCode", 0, 0, 0, 0)
  Dim TeamKeyword
  TeamKeyword = ""
  Dim Kontierungsnr
  Kontierungsnr = hlObj.GetValue("ProcurementDetail.AllocationNumber", 0, 0, 0, 0)
  Dim Kontierungstyp
  Kontierungstyp = hlObj.GetValue("ProcurementDetail.AllocationType", 0, 0, 0, 0)
  Dim Beschreibung
  Beschreibung = ""
  Beschreibung = "CI ist auf Status 'Elimiert' gesetzt worden. Die CI-Nummmer steht im Betreff. Der Eliminierungsgrund lautet: " & ElimierungsgrundDE
  Beschreibung = Beschreibung & CHR(13) & CHR(10) & "The CI-Status is set to Eliminated. The CI-Number is displayed in the subject of the incident. The elimination reason is: " & ElimierungsgrundEN
  Beschreibung = Beschreibung & CHR(13) & CHR(10) & "Kontierungsnummer: " & Kontierungsnr
  Beschreibung = Beschreibung & CHR(13) & CHR(10) & "Kontierungstyp: " & Kontierungstyp
  Beschreibung = Beschreibung & CHR(13) & CHR(10) & "Allocationnumber: " & Kontierungsnr
  Beschreibung = Beschreibung & CHR(13) & CHR(10) & "Allocationtype: " & Kontierungstyp


  SELECT CASE Buchungskreis

    CASE "107"
      TeamKeyword = "KOControllingDitzingen"
    CASE "110"
      TeamKeyword = "KOControllingDitzingen"
    CASE "111"
      TeamKeyword = "KOControllingDitzingen"
    CASE "114"
      TeamKeyword = "KOControllingDitzingen"
    CASE "122"
      TeamKeyword = "KOControllingDitzingen"
    CASE "146"
      TeamKeyword = "KOControllingDitzingen"
    CASE "222"
      TeamKeyword = "KOControllingGruesch"
    CASE "223"
      TeamKeyword = "KOControllingGruesch"
    CASE "225"
      TeamKeyword = "KOControllingGruesch"
    CASE "314"
      TeamKeyword = "KOControllingPasching"
    CASE "231"
      TeamKeyword = "KOControllingFarmington"
    CASE "237"
      TeamKeyword = "KOControllingCranbury"
  END SELECT

  '///////////////////////////////////////////////////////////
  Dim cinummer
  cinummer = hlObj.GetValue("TrumpfAssetGeneral.CINumber", 0, 0, 0, 0)
  Dim increqsubject
  increqsubject = "Eliminierung/Elimination: " & cinummer & " Internal helpLine-ID: " & hlObj.GetID

  ' VBScript source code
  Dim xmldoc
  Set xmldoc = CreateObject("msxml2.DomDocument")

  'create root element
  Dim nodeData
  Set nodeData = xmldoc.appendChild(xmldoc.createElement("Data"))
  Dim nodeObjects
  Set nodeObjects = nodeData.appendChild(xmldoc.createElement("Objects"))

  '//// obj1: IncidentRequest///////////////////////////////////////////////////////

  Dim nodeObject
  Set nodeObject = nodeObjects.appendChild(xmldoc.createElement(objDefname))
  Dim attAliasName
  Set attAliasName = xmldoc.createAttribute("aliasname")
  attAliasName.Text = aliasname1
  nodeObject.Attributes.setNamedItem attAliasName
  Dim nodeAttributes
  Set nodeAttributes = nodeObject.appendChild(xmldoc.createElement("Attributes"))
  Dim nodeServiceUnits
  Set nodeServiceUnits = nodeObject.appendChild(xmldoc.createElement("ServiceUnits"))
  Dim nodeServiceUnit
  Set nodeServiceUnit = nodeServiceUnits.appendChild(xmldoc.createElement("ServiceUnit"))
  Dim attAliasNameSU
  Set attAliasNameSU = xmldoc.createAttribute("aliasname")
  attAliasNameSU.Text = aliasnameSU
  nodeServiceUnit.Attributes.setNamedItem attAliasNameSU
  '//////////////// HLOBJECT.ID
  'Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
  Call AppendNode(hlContext, xmldoc, nodeAttributes, True, "CaseGeneral.Subject", increqsubject)
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "CaseDescription.DescriptionText", Beschreibung)
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "Keywords.KeywordOrga", TeamKeyword)
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "CaseGeneral.CompanyCode", Buchungskreis1)
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "Keywords.Keyword", "KWStdSWhelplineInterfaceAM")
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "IncidentAttribute.IncidentStatus", "IncidentStatusNew")
  Call AppendNode(hlContext, xmldoc, nodeAttributes, False, "IncidentAttribute.RequestType", "RequestTypeService")
  Call AppendNode(hlContext, xmldoc, nodeServiceUnit, True, "IncidentSUAttribute.IncidentOperation", "IncidentOperation")

  '//// obj2: Product///////////////////////////////////////////////////////
  Dim nodeObject2
  Set nodeObject2 = nodeObjects.appendChild(xmldoc.createElement(hlObj.GetType()))
  Dim attAliasName2
  Set attAliasName2 = xmldoc.createAttribute("aliasname")
  attAliasName2.Text = aliasname2
  nodeObject2.Attributes.setNamedItem attAliasName2
  Dim nodeAttributes2
  Set nodeAttributes2 = nodeObject2.appendChild(xmldoc.createElement("Attributes"))
  Call AppendNode(hlContext, xmldoc, nodeAttributes2, True, "TrumpfAssetGeneral.CINumber", cinummer)


  '//// Relations///////////////////////////////////////////////////////
  Dim nodeRelations
  Set nodeRelations = nodeData.appendChild(xmldoc.createElement("Relations"))
  Dim nodeProduct2Case
  Set nodeProduct2Case = nodeRelations.appendChild(xmldoc.createElement("Product2Case"))

  Call AppendTextNode(hlContext, xmldoc, nodeProduct2Case, "Parent", aliasnameSU)
  Call AppendTextNode(hlContext, xmldoc, nodeProduct2Case, "Child", aliasname2)

  ' Save to File
  Dim Filename
  Filename = "\\srvditz1\pi_intern\helpline\helpline_in\c11\" & objDefname & "_" & hlObj.GetID & ".xml"

  xmldoc.Save(Filename)

End Sub

Public Sub AppendNode(ByRef hlContext, ByRef xmldoc, ByRef nodeObject, ByRef iskey, ByRef key, ByRef value)
  Dim valueNode
  Set valueNode = xmldoc.createElement(key)
  Dim cdata
  Set cdata = xmldoc.createCDATASection(value)
  valueNode.appendChild(cdata)
  nodeObject.appendChild(valueNode)

  Dim attIsKey
  Set attIsKey = xmldoc.createAttribute("iskey")
  IF (iskey) THEN
    attIsKey.Text = "true"
  ELSE
    attIsKey.Text = "false"
  END IF
  valueNode.Attributes.setNamedItem attIsKey

End Sub

Public Sub AppendTextNode(ByRef hlContext, ByRef xmldoc, ByRef nodeObject, ByRef key, ByRef value)
  Dim valueNode
  Set valueNode = xmldoc.createElement(key)
  nodeObject.appendChild(valueNode)

  valueNode.Text = value
End Sub

Public Function DBConnectionString(ByRef hlContext)
  Const Skrypton.LegacyParser.Tokens.Basic.NameToken:DBConnection = Skrypton.LegacyParser.Tokens.Basic.StringToken:Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2

End Function
