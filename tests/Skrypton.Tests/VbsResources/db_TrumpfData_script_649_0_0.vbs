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
Function hlITIL2()
	Set hlITIL2 = CreateObject("hlStartITIL2.Global")
	hlITIL2.SelfCheck hlContext
End Function
'----------------------------------------------------------------------------------------------------------
'Globale Konstanten fuer freie Assoziationsdefinitionen
Const HLASC_SoftwareLicenseFolderView = 110941
Const HLASC_SoftwareLicenseGroupView = 110944
'----------------------------------------------------------------------------------------------------------
'Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
Sub Trace(hlContext,text)
 	hlContext.trace 1, text
End Sub
'----------------------------------------------------------------------------------------------------------
'Funktion InfoMail
'Zum Aufrufen aus EBL-Skripten von Vorgaengen
Sub InfoMail(hlContext, hlCase, Subject, MailSender, Receiver, CC, body, SendAttachments) 

	Dim Email
	Set Email = hlContext.CreateMail
	
	'Falls der Parameter <SendAttachmnets> beim Aufruf "1" ist, werden Anhaenge mitversandt
	If CBool(SendAttachments)=True Then
		Dim AttachIDs
		Dim AttachID
		Dim Attachment : Set Attachment = Nothing
		AttachIDs = hlCase.GetAttachmentKeys("HLOBJECTINFO.ATTACHMENT",0)	
		For Each AttachID in AttachIDs
			Set Attachment = hlCase.GetAttachment("HLOBJECTINFO.ATTACHMENT", AttachID, 0)
			If Attachment.Size > 0 Then
				Dim MailAttachment : Set MailAttachment = Nothing
				Set MailAttachment = Email.AddAttachment
				MailAttachment.name = Attachment.name
				MailAttachment.data = Attachment.data
			End If
		Next
	End If
	
	If MailSender <> "" Then
		Email.SenderMail = MailSender
	End If				
	Email.To = Receiver  
	Email.Subject = Subject
	Email.Body = body
	If CC <> "" Then
		Email.CC = CC
	End If
	Call hlContext.SendRequestMail(Email)
End Sub

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
Sub CreateSubject(hlContext,Survey,hlCaller)
	Dim language
	language = hlCaller.GetValue("PersonGeneral.Language",0,0,0,0)
	If language = "LanguageGerman" Then
		Call	Survey.SetValue("CaseGeneral.Subject",0,0,0,"Umfrage zur Service-Leistung ihres Support-Teams")
	Else
		Call Survey.SetValue("CaseGeneral.Subject",0,0,0,"Survey about the Service-Quality from your Support-Team")
	End If	
End Sub
'----------------------------------------------------------------------------------------------------------
Sub InviteSurveyEmail(hlContext,hlCase,hlCaller)
	'Email an den Anfrager eines Survey-Vorgangs, um diesen zur Teilnahme an der
	'Umfrage aufzufordern. 
	'Email to Requester of a Survey-Case to invite him to take part on the survey
	Dim SUIDx : SUIDx = hlITIL2.GetLastSUIdx(hlCase,hlContext)
	Dim MailRequest : MailRequest = hlCase.GetValue("CaseGeneral.DefaultNotification",0,0,0,0)
	If MailRequest = "DefaultNotificationEmail" And SUIDx = 1 Then
		Dim strCRLF : strCRLF = CHR(13) & CHR(10)
		Dim Creationdate, Datum, subject, body
		Dim refnumber : refnumber = hlcase.GetValue("CASEINFO.REFERENCENUMBER",0,0,0,0) 
		Dim portallink : portallink = "http://localhost/helplineportal/"
		Dim surname : surname = hlCaller.GetValue("PersonGeneral.PersonSurname",0,0,0,0)
		Dim letteraddress : letteraddress = hlCaller.GetValue("PersonGeneral.ShortLetterAddress",0,0,0,0)
		Dim Anrede : Anrede = "Sehr geehrte Damen und Herren,"
		Dim PersonAddress : PersonAddress = "Dear Mrs./Ms. or Mr.,"	
		Dim language
		language = hlCaller.GetValue("PersonGeneral.Language",0,0,0,0)
	
		If language = "LanguageGerman" Then
			If letteraddress = "" Then
				letteraddress = "Herr/Frau"
			End If
			Anrede = "Sehr geehrte(r) " & CStr(letteraddress) & " " & CStr(surname) & ","
	
			'Hier wird die Betreffzeile erstellt
			'The subject field is entered here
			Creationdate = hlcase.GetValue("HLOBJECTINFO.CREATIONTIME",7,0,0,0)
			Datum = Mid(Creationdate,1,10)
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
			body = body & "Dort finden Sie das Umfrage-Formular mit der Nummer " & refnumber &". "
			body = body & strCRLF & strCRLF		
			body = body & "Wir freuen uns sehr, wenn Sie sich die Zeit nehmen, die Fragen zu beantworten."
			body = body & strCRLF & strCRLF
			body = body & "Wir bedanken uns fuer Ihre Unterstuetzung!"
			body = body & strCRLF & strCRLF
			body = body & strCRLF & "Mit freundlichen Gruessen"
			body = body & strCRLF & strCRLF
			body = body & "Ihr Support Team"
		Else	
			If letteraddress = "" Then
				letteraddress = "Mrs./Ms./Mr."
			End If
			PersonAddress = "Dear " & CStr(letteraddress) + " " & CStr(surname) & ","
		
			'Hier wird die Betreffzeile erstellt
			'The subject field is entered here
			Creationdate = hlcase.GetValue("HLOBJECTINFO.CREATIONTIME",7,0,0,0)
			Datum = Mid(Creationdate,1,10)
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
			body = body & "Then klick 'Survey' in the menue 'Your Requests'. "& strCRLF
			body = body & "There, you will find the Questionnaire with the reference number " & refnumber &". "
			body = body & strCRLF & strCRLF		
			body = body & "It would be nice, if you invest your time to response the questions."
			body = body & strCRLF & strCRLF
			body = body & "We thank you for your assistance!"
			body = body & strCRLF & strCRLF
			body = body & strCRLF & "With best regards"
			body = body & strCRLF & strCRLF
			body = body & "Yours Support Team"
		End If
 
	 	Dim Email
		Set Email = hlContext.CreateMail
	
		'Ermittle die Emailadresse des Anfragers
		'Detect email adress of requester
		Dim Emailadress : Emailadress = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
		If Emailadress = "" Then
			Emailadress= "Username@yourcompany.com"
			subject = "Diese EMail konnte nicht zugestellt werden"
			body = "Die Mail fuer die Anfragenummer "
			body = body & hlcase.GetValue("CASEINFO.REFERENCENUMBER",0,0,0,0)
			body = body & " konnte wegen einer fehlenden E-Mail Adresse nicht zugestellt werden."			
		End If
		Email.To = Emailadress
		Email.Subject = subject
		Email.Body = body		
		Call hlContext.SendRequestMail(Email)	
	End If
End Sub
'----------------------------------------------------------------------------------------------------------
'Diese Funktion steuert den SystemTask wenn dieser im Vorgangstyp Task konfiguriert wurde.
'This function controls a SystemTask if it had been configured within the casetype Task.
Sub MyTask1(hlContext)
	Dim hlObj : Set hlObj = hlContext.GetCurrentObject()
	Dim lcid : lcid = 0
	lcid = hlContext.GetLocaleID
	Dim LangID : LangID = 0
	LangID = hlContext.LangIDFromLCID(lcid)
	
	'Gesetzte Daten aus dem aktuellen Task auslesen, diese werden dem zu erzeugenden Systemtask mitgegeben.
	'Read setted data of current task and take them into the created Systemtask.
	Dim Priority : Priority = hlObj.GetValue("CaseClassificationAttribute.Priority",0,0,0,0)
	Dim TaskType : TaskType = hlObj.GetValue("TaskGeneral.TaskType",0,0,0,0)
	Dim Subject : Subject = hlObj.GetValue("TaskGeneral.Subject",0,0,0,0)
	Dim Description : Description = hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0)
	Dim ExOperation : ExOperation = hlObj.GetValue("CaseDiagnosis.DiagnosisText",0,0,0,0)
	Dim AssignedGroup : AssignedGroup = hlObj.GetValue("CaseSpecialRouting.AssignedGroup",0,0,0,0)
	Dim AssignedPerson : AssignedPerson = hlObj.GetValue("CaseSpecialRouting.AssignedPerson",0,0,0,0)
	Dim Team : Team = hlObj.GetValue("Keywords.KeywordOrga",0,0,0,0)	
	Dim newTask	
	Set newTask = hlContext.createobject("Task")
	
	newTask.SetValue "CaseClassificationAttribute.Priority", 0, 0, 0, Priority
	newTask.SetValue "TaskGeneral.TaskType", 0, 0, 0, TaskType
	newTask.SetValue "TaskGeneral.Subject", 0, 0, 0, Subject
	newTask.SetValue "Keywords.KeywordOrga", 0, 0, 0, Team		
	
	Dim hasContent
	hasContent = hlObj.HasContent("TaskDesignWorkflow.TaskWorkflow_CA", 0, 0)
	If hasContent <> 0 Then
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
		For Each contentID in contentIDs
			assignedGroupWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup", 0, contentID, 0, 0)
			assignedPersonWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson", 0, contentID, 0, 0)
			descriptionWF = hlObj.GetValue("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText", 0, contentID, 0, 0)
			newContentID = hlObj.GenerateContentID
			newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup", 0, newContentID, 0, assignedGroupWF
			newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson", 0, newContentID, 0, assignedPersonWF
			newTask.SetValue "TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText", 0, newContentID, 0, descriptionWF
		Next
	End If
	
	
	Dim Assets 
	Dim Asset
	Assets = hlObj.GetItemsEx(0,0,131)
		For Each Asset in Assets
			Call	newTask.AddItemEx(0,Asset,0,131)
		Next
	Dim RefNumber : RefNumber = hlObj.GetValue("CASEINFO.REFERENCENUMBER",0,0,0,0)
	If LangID = 7 Then
		Description = Description & vbNewLine & vbNewLine & "[Diese Aufgabe wurde automatisch durch den Systemtask mit der Bezugsnummer '" & RefNumber & " erstellt.]"
	Else
		Description = Description & vbNewLine & vbNewLine & "[This Task was created automatically by Systemtask with the Reference Number '" & RefNumber & "'.]"
	End If
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
Sub CreateSystemTaskDefbyDay(SysTaskBeginnDate, SysTaskEndDate, NoEndDate, NumberOfDays, taskDefname, recurrenceEndType)
	Dim hlObj : Set hlObj = hlContext.GetCurrentObject()
	Dim hlSystemTask : Set hlSystemTask= hlContext.CreateSystemTask(0)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Dim systemTaskDefinitionName : systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME",0,0,0,0)
	Dim scriptCode : scriptCode = "MyTask1"
	Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE",0,0,0,SysTaskBeginnDate)
	'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
	'Check which option had been chosen in duration of the SystemTask.
	Dim	newTaskEndTime
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
	If recurrenceEndType = "0" Then
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
		recurrenceEndType = "1"
	Else
		If recurrenceEndType = "2" Then
		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
		recurrenceEndType = "1"
		End If
	End If
	'Neu - Ende

	Call hlSystemTask.SetValue("SYSTASKINFO.ENDTYPE",0,0,0,recurrenceEndType)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE",0,0,0,scriptCode)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL",0,0,0,NumberOfDays)							
	Call hlContext.SaveSystemTask(hlSystemTask)
	Dim hlSystemTaskDefinitionObj : Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
	Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub

'Entfernt einen vorhandenen SystemTask.
'Remove an existing SystemTask.
Sub DeleteSystemTask(hlContext, hlObj, hlSystemTask, taskname)
  Call hlContext.RemoveSystemTask(hlSystemTask)
End Sub

'Festlegung der Definitionen eines SystemTasks pro Woche.
'Determining of definitions of a SystemTask by week.
Sub CreateSystemTaskDefbyWeek(SysTaskBeginnDate, SysTaskEndDate, NoEndDate, NumberOfWeeks, MondayFlag, TuesdayFlag, WednesdayFlag, ThursdayFlag, FridayFlag, SaturdayFlag, SundayFlag, taskDefname, recurrencedaymask, recurrenceEndType)
	Dim hlObj : Set hlObj = hlContext.GetCurrentObject()
	Dim hlSystemTask : Set hlSystemTask = hlContext.CreateSystemTask(0)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Dim systemTaskDefinitionName : systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME",0,0,0,0)
	Dim scriptCode : scriptCode = "MyTask1"
	
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
	If recurrenceEndType = "0" Then
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
		recurrenceEndType = "1"
	Else
		If recurrenceEndType = "2" Then
		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
		recurrenceEndType = "1"
		End If
	End If
	'Neu - Ende
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE",0,0,0,scriptCode)
	Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE",0,0,0,SysTaskBeginnDate)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE",0,0,0,recurrenceEndType)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL",0,0,0,NumberOfWeeks)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.DAYMASK",0,0,0,recurrencedaymask)
	
  	Call hlContext.SaveSystemTask(hlSystemTask)
	Dim hlSystemTaskDefinitionObj : Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
	Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub  

'Festlegung der Definitionen eines SystemTasks pro Monat.
'Determining of definitions of a SystemTask by month.
Sub CreateSystemTaskDefbyMonth(SysTaskBeginnDate, SysTaskEndDate, NoEndDate, DayOfMonth, NumberOfMonths, taskDefname, recurrenceEndType)
	Dim hlObj : Set hlObj = hlContext.GetCurrentObject()
	Dim hlSystemTask : Set hlSystemTask = hlContext.CreateSystemTask(0)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Dim systemTaskDefinitionName : systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME",0,0,0,0)
	Dim scriptCode : scriptCode = "MyTask1"
	
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
	If recurrenceEndType = "0" Then
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
		recurrenceEndType = "1"
	Else
		If recurrenceEndType = "2" Then
		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
		recurrenceEndType = "1"
		End If
	End If
	'Neu - Ende
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE",0,0,0,recurrenceEndType)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.DAYOFMONTH",0,0,0,DayOfMonth)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INSTANCE",0,0,0,"0")
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL",0,0,0,NumberOfMonths)
	Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE",0,0,0,SysTaskBeginnDate)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE",0,0,0,scriptCode)
	Call hlContext.SaveSystemTask(hlSystemTask)
	Dim hlSystemTaskDefinitionObj : Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
	Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
End Sub

'Sub fuehrt den SystemTask einmalig aus.
'Sub execute SystemTask one-time.
Sub CreateOneTimeSystemTask(OneTimeTask, SysTaskEndDate, SysTaskBeginnDate, taskDefname)
	Dim hlObj : Set hlObj = hlContext.GetCurrentObject()
	Dim hlSystemTask : Set hlSystemTask = hlContext.CreateSystemTask("0")
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Dim scriptCode : scriptCode = "MyTask1"

	'Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
	'Check which option had been chosen in duration of the SystemTask.
	Dim systemTaskDefinitionName : systemTaskDefinitionName = hlSystemTask.GetValue("SYSTASKINFO.DEFNAME",0,0,0,0)
	Call hlSystemTask.SetValue("SYSTASKINFO.DEFNAME",0,0,0,taskDefname)
	Call hlSystemTask.SetValue("SYSTASKINFO.STARTDATE",0,0,0,SysTaskBeginnDate)
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,"09.09.2099 09:09:09")
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDTYPE",0,0,0,1) 	
	Call hlSystemTask.SetValue("SYSTASKINFO.ENDCOUNT",0,0,0,1)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.TYPE",0,0,0,0)
	Call hlSystemTask.SetValue("SYSTASKINFO.RECURRENCE.INTERVAL",0,0,0,1)
 	Call hlSystemTask.SetValue("SYSTASKINFO.SCRIPTCODE",0,0,0,scriptCode)
 	
	'Wenn kein Datum angegebene wurde, muss eine Fehlermeldung angezeigt werden.
	'If no date was entered, show an error message.
	If SysTaskBeginnDate = "" Then
		errCode = "#ERR_TSKMNT_002"
	End If
	Call hlContext.SaveSystemTask(hlSystemTask)
	Dim hlSystemTaskDefinitionObj : Set hlSystemTaskDefinitionObj = hlContext.GetSystemTask(systemTaskDefinitionName)
	Call hlObj.AddSystemtask(hlSystemTaskDefinitionObj)
	Call hlObj.HasSystemTask(systemTaskDefinitionName)
End Sub
'----------------------------------------------------------------------------------------------------------
Function IsValidObject(obj)
	IsValidObject = (IsObject(obj) And ( Not (obj Is Nothing)) )
End Function

'XML-Export Neuanlage

Sub ExportObject(hlContext,hlObj)
    Dim objDefname : objDefname = hlObj.GetType()
    Dim aliasname : aliasname = "NewCI"&objDefname 
    Dim NewChangeObj : NewChangeObj = hlObj.GetValue("TrumpfAssetGeneral.DataToSAPAMChange",0,0,0,0)
    If NewChangeObj = "0" Or NewChangeObj = "" Then
			aliasname = aliasname 
		Else
			aliasname ="ChangedCI"&objDefname
		End If

    ' VBScript source code
    Dim xmldoc
    Set xmldoc = CreateObject("msxml2.DomDocument")

    'create root element
    Dim nodeData : Set nodeData = xmldoc.appendChild(xmldoc.createElement("Data"))
    Dim nodeObjects : Set nodeObjects = nodeData.appendChild(xmldoc.createElement("Objects"))
    Dim nodeObject : Set nodeObject = nodeObjects.appendChild(xmldoc.createElement(objDefname))
    Dim attAliasName : Set attAliasName = xmldoc.createAttribute("aliasname")
    attAliasName.Text = aliasname
    nodeObject.Attributes.setNamedItem attAliasName
    Dim nodeAttributes : Set nodeAttributes = nodeObject.appendChild(xmldoc.createElement("Attributes"))
    Dim nodeRelations : Set nodeRelations = nodeData.appendChild(xmldoc.createElement("Relations"))
    '///////////////////////////////////////////////////////////

    '//////////////// HLOBJECT.ID
    'Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) ' 
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "AssetGeneral.AssetName", hlObj.GetValue("AssetGeneral.AssetName", 0,0,0,0))' hlObj.GetValue("AssetGeneral.AssetName", 0,0,0,0)
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "AccountingDetail.CostCenter", hlObj.GetValue("AccountingDetail.CostCenter", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "AssetGeneral.Serialnumber", hlObj.GetValue("AssetGeneral.Serialnumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "ProcurementDetail.AllocationNumber", hlObj.GetValue("ProcurementDetail.AllocationNumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "ProcurementDetail.AllocationType", hlObj.GetValue("ProcurementDetail.AllocationType", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "ProcurementDetail.OrderNumber", hlObj.GetValue("ProcurementDetail.OrderNumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "ProcurementDetail.OrderPosition", hlObj.GetValue("ProcurementDetail.OrderPosition", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "ProcurementDetail.VendorNumber", hlObj.GetValue("ProcurementDetail.VendorNumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "TrumpfAssetGeneral.CINumber", hlObj.GetValue("TrumpfAssetGeneral.CINumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.CompanyCode", hlObj.GetValue("TrumpfAssetGeneral.CompanyCode", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.InvestmentNumber", hlObj.GetValue("TrumpfAssetGeneral.InvestmentNumber", 0,0,0,0))
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.Manufacturer", hlObj.GetValue("TrumpfAssetGeneral.Manufacturer", 0,0,0,0))                            
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "TrumpfAssetGeneral.SAPCostCenter", hlObj.GetValue("TrumpfAssetGeneral.SAPCostCenter", 0,0,0,0))                            

    ' Save to File
	Dim Filename
	If NewChangeObj = "0" Or NewChangeObj = "" Then
		Filename = "\\srvditz1\pi_intern\helpline\HELPLINE_out\c11\hlnew" & objDefname &"_" & hlObj.GetID & ".xml"
	Else
		Filename = "\\srvditz1\pi_intern\helpline\HELPLINE_out\c11\hlchange" & objDefname &"_" & hlObj.GetID & ".xml"
	End If	
	xmldoc.Save(Filename)

End Sub

Sub AppendNode(hlContext,xmldoc, nodeObject, iskey, key, value)
    Dim valueNode
    Set valueNode= xmldoc.createElement(key)
    Dim cdata
    Set cdata = xmldoc.createCDATASection(value)
    valueNode.appendChild(cdata)
    nodeObject.appendChild(valueNode)
    
    Dim attIsKey : Set attIsKey = xmldoc.createAttribute("iskey")
    If (iskey) Then 
	attIsKey.Text = "true" 
    Else
	attIsKey.Text = "false"
    End If
    valueNode.Attributes.setNamedItem attIsKey
    
End Sub

'XML-Export Incident wegen Eleminierung

Sub ExportObjectIncident(hlContext,hlObj)
    Dim objDefname : objDefname = "IncidentRequest"
    Dim aliasname1 : aliasname1 = "obj1"
    Dim aliasname2 : aliasname2 = "obj2"
    Dim aliasnameSU : aliasnameSU = "objSU"
    Dim ElimierungsgrundDE : ElimierungsgrundDE = hlObj.GetValue("TrumpfAssetStatus.CISubStatus",7,0,0,0) 
    Dim ElimierungsgrundEN : ElimierungsgrundEN = hlObj.GetValue("TrumpfAssetStatus.CISubStatus",9,0,0,0)
    Dim Buchungskreis : Buchungskreis = hlObj.GetValue("TrumpfAssetGeneral.CompanyCode",0,0,0,0)
    Dim Buchungskreis1 : Buchungskreis1 = hlObj.GetValue("TrumpfAssetGeneral.CompanyCode",0,0,0,0)
    Dim TeamKeyword : TeamKeyword = ""
    Dim Kontierungsnr : Kontierungsnr = hlObj.GetValue("ProcurementDetail.AllocationNumber",0,0,0,0)
    Dim Kontierungstyp : Kontierungstyp = hlObj.GetValue("ProcurementDetail.AllocationType",0,0,0,0)
    Dim Beschreibung : Beschreibung = ""
    Beschreibung = "CI ist auf Status 'Elimiert' gesetzt worden. Die CI-Nummmer steht im Betreff. Der Eliminierungsgrund lautet: " &ElimierungsgrundDE
    Beschreibung = Beschreibung&CHR(13)&CHR(10)&"The CI-Status is set to Eliminated. The CI-Number is displayed in the subject of the incident. The elimination reason is: "&ElimierungsgrundEN
    Beschreibung = Beschreibung&CHR(13)&CHR(10)&"Kontierungsnummer: "&Kontierungsnr
    Beschreibung = Beschreibung&CHR(13)&CHR(10)&"Kontierungstyp: "&Kontierungstyp 
    Beschreibung = Beschreibung&CHR(13)&CHR(10)&"Allocationnumber: "&Kontierungsnr
    Beschreibung = Beschreibung&CHR(13)&CHR(10)&"Allocationtype: "&Kontierungstyp    
     
    
    Select Case Buchungskreis
    		Case "107"
    			TeamKeyword = "KOControllingDitzingen"
    		Case "110"
    			TeamKeyword = "KOControllingDitzingen"
    		Case "111"
    			TeamKeyword = "KOControllingDitzingen"
    		Case "114"
    			TeamKeyword = "KOControllingDitzingen"
		Case "122"
    			TeamKeyword = "KOControllingDitzingen"
    		Case "146"
    			TeamKeyword = "KOControllingDitzingen"
    		Case "222"
    			TeamKeyword = "KOControllingGruesch"
    		Case "223"
    			TeamKeyword = "KOControllingGruesch"
    		Case "225"
    			TeamKeyword = "KOControllingGruesch"
		Case "314"
    			TeamKeyword = "KOControllingPasching"
		Case "231"
    			TeamKeyword = "KOControllingFarmington"
		Case "237"
    			TeamKeyword = "KOControllingCranbury"
    End Select
    
 '///////////////////////////////////////////////////////////
    Dim cinummer : cinummer = hlObj.GetValue("TrumpfAssetGeneral.CINumber", 0,0,0,0)
    Dim increqsubject : increqsubject = "Eliminierung/Elimination: "&cinummer &" Internal helpLine-ID: "&hlObj.GetID    

    ' VBScript source code
    Dim xmldoc
    Set xmldoc = CreateObject("msxml2.DomDocument")

    'create root element
    Dim nodeData : Set nodeData = xmldoc.appendChild(xmldoc.createElement("Data"))
    Dim nodeObjects : Set nodeObjects = nodeData.appendChild(xmldoc.createElement("Objects"))

    '//// obj1: IncidentRequest///////////////////////////////////////////////////////
        
    Dim nodeObject : Set nodeObject = nodeObjects.appendChild(xmldoc.createElement(objDefname))
    Dim attAliasName : Set attAliasName = xmldoc.createAttribute("aliasname")
    attAliasName.Text = aliasname1
    nodeObject.Attributes.setNamedItem attAliasName
    Dim nodeAttributes : Set nodeAttributes = nodeObject.appendChild(xmldoc.createElement("Attributes"))
    Dim nodeServiceUnits : Set nodeServiceUnits = nodeObject.appendChild(xmldoc.createElement("ServiceUnits"))
    Dim nodeServiceUnit : Set nodeServiceUnit = nodeServiceUnits.appendChild(xmldoc.createElement("ServiceUnit"))
    Dim attAliasNameSU : Set attAliasNameSU= xmldoc.createAttribute("aliasname")
    attAliasNameSU.Text = aliasnameSU
    nodeServiceUnit.Attributes.setNamedItem attAliasNameSU
    '//////////////// HLOBJECT.ID
    'Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) ' 
    Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "CaseGeneral.Subject", increqsubject ) 
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "CaseDescription.DescriptionText", Beschreibung)
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "Keywords.KeywordOrga", TeamKeyword)
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "CaseGeneral.CompanyCode", Buchungskreis1)
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "Keywords.Keyword", "KWStdSWhelplineInterfaceAM")
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "IncidentAttribute.IncidentStatus", "IncidentStatusNew")                            
    Call AppendNode(hlContext,xmldoc, nodeAttributes, False, "IncidentAttribute.RequestType", "RequestTypeService")
    Call AppendNode(hlContext,xmldoc, nodeServiceUnit, True, "IncidentSUAttribute.IncidentOperation","IncidentOperation") 
    
    '//// obj2: Product///////////////////////////////////////////////////////
    Dim nodeObject2 : Set nodeObject2 = nodeObjects.appendChild(xmldoc.createElement(hlObj.GetType()))
    Dim attAliasName2 : Set attAliasName2 = xmldoc.createAttribute("aliasname")
    attAliasName2.Text = aliasname2
    nodeObject2.Attributes.setNamedItem attAliasName2
    Dim nodeAttributes2 : Set nodeAttributes2 = nodeObject2.appendChild(xmldoc.createElement("Attributes"))
    Call AppendNode(hlContext,xmldoc, nodeAttributes2, True, "TrumpfAssetGeneral.CINumber", cinummer) 
    
    
    '//// Relations///////////////////////////////////////////////////////
    Dim nodeRelations : Set nodeRelations = nodeData.appendChild(xmldoc.createElement("Relations"))
    Dim nodeProduct2Case: Set nodeProduct2Case = nodeRelations.appendChild(xmldoc.createElement("Product2Case"))
    
    Call AppendTextNode(hlContext,xmldoc, nodeProduct2Case, "Parent", aliasnameSU)    
    Call AppendTextNode(hlContext,xmldoc, nodeProduct2Case, "Child",  aliasname2)

    ' Save to File
	Dim Filename : Filename = "\\srvditz1\pi_intern\helpline\helpline_in\c11\" & objDefname &"_" & hlObj.GetID & ".xml"
	
	xmldoc.Save(Filename)

End Sub

Sub AppendNode(hlContext,xmldoc, nodeObject, iskey, key, value)
    Dim valueNode
    Set valueNode= xmldoc.createElement(key)
    Dim cdata
    Set cdata = xmldoc.createCDATASection(value)
    valueNode.appendChild(cdata)
    nodeObject.appendChild(valueNode)
    
    Dim attIsKey : Set attIsKey = xmldoc.createAttribute("iskey")
    If (iskey) Then 
	attIsKey.Text = "true" 
    Else
	attIsKey.Text = "false"
    End If
    valueNode.Attributes.setNamedItem attIsKey
    
End Sub

Sub AppendTextNode(hlContext,xmldoc, nodeObject, key, value)
    Dim valueNode
    Set valueNode= xmldoc.createElement(key)
    nodeObject.appendChild(valueNode)
    
    valueNode.Text = value
End Sub

Function DBConnectionString (hlContext)
	Const DBConnection = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2"
End Function 
