Function GetCharmURL()
        GetCharmURL = "http://helpline.wattens.swarovski.com/ManagementPortal/Charm/Create/"
End Function
Function GetOpenCharmURL()
        GetOpenCharmURL = "https://p92.internal.swarovski.com/sap/public/bc/workflow/shortcut?sysid=P92&client=010&transaction=*ZSCHARM_SHORTCUT&param=PA_OBJ="
End Function

Function GetCatsURL()
        GetCatsURL = "http://helpline.wattens.swarovski.com/ManagementPortal/Cats/Book/"
End Function


'Standard Change Specific 

Function StandardChangeOnLoad()
	Dim Prozent
	Dim Items
	Dim ControlsCA
	Dim Control
	Dim Label
	Dim count
	Dim Atta
	Dim AttCount
	Dim AgentID
	Dim Person
	Dim GroupIDs

	'Get Agent
	AgentID = hlSession.GetAgentID()
	Set Person = GetPersonForAgent(AgentID)
	If hlObj.IsNew then ButtonCancel.ShowControl = 3
	If model.IsInWeb = true then ButtonDocumentation.ShowControl = 3

	If hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0) <> "" then
		LabelReservedBy.Text = "Case reserved by: " & GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
	Else
		LabelReservedBy.Text = ""
	End If

	'Set ToolTips

	TextBoxStandardChangeCategoryText.ToolTip = TextBoxStandardChangeCategoryText.Text
	TextBoxName.ToolTip = TextBoxName.Text
	TextBoxGroupName.ToolTip = TextBoxGroupName.Text
	ComboBoxAssignee.ToolTip = ComboBoxAssignee.Text

	'Disable Assignee Combobox
	If SearchButtonGroup.GetSearchState <> 3 Then
		ComboBoxAssignee.Disabled = true
	Else
		ComboBoxAssignee.Search.SearchCondition = "PersonDisplayHelper.GroupIDText Like ""*" & GroupID & "*"" AND ConfigurationAttribute.AvailabilityStatus = ""AvailabilityStatusAvailable"""
	End If

	'Set Attachment Count
	Atta = hlObj.GetAttachmentKeys("HLOBJECTINFO.ATTACHMENT",count)
	AttCount = 0
	For Each Att In Atta
		AttCount = AttCount +1
	next
	TabPageAttachment.Caption = "Attachment (" & AttCount & ")"

	'Additional Recipient Count befüllen
	TabPageAdditionalRecipients.Caption = "Additional Recipients (" & hlObj.GetItemCount(&H00000, "AdditionalRecipient2Case") & ")"

	'Parent/Child Count befüllen
	TabPageParentChild.Caption = "Parent (" & hlObj.GetItemCount(&H10000, "StandardChangeRecord2StandardChangeRecord") & ")/Child (" & hlObj.GetItemCount(&H00000, "StandardChangeRecord2StandardChangeRecord") &  ")"

	'Special Handling
	If hlCaller.GetValue("PersonSpecific.SpecialHandling",1033,0,0,0) <> "" Then 
		GroupBoxRequesterSearch.Caption = "Affected End User - " & hlCaller.GetValue("PersonSpecific.SpecialHandling",1033,0,0,0)
		hlObj.SetValue "StandardChangeRecordSpecific.SpecialHandling", 0, 0, 0, hlCaller.GetValue("PersonSpecific.SpecialHandling",0,0,0,0)
	End If

	'Disable Buttons

	count = hlObj.GetSvcUnitCount

	If hlObj.IsReadOnly("CaseDescription.Description.RAWTEXT",count)=1 Then 
		ButtonCloseTask.Disabled = true
		ButtonWaitingFor.Disabled = true
		ButtonCancel.Disabled = true
	Else
		ButtonCloseTask.Disabled = false
		ButtonWaitingFor.Disabled = false
		ButtonCancel.Disabled = false
	End If

	'Progressbar

	Prozent = hlObj.GetValue("StandardChangeRecordSpecific.Completed",1033,0,0,0)
	If Prozent > 0 Then
		For i = 10 To 100 Step 10
			if i <= Cint(Prozent) then
				Label = "Label" & i & "P"
				model.GetControlFromID(Label).BackColor="#1b709f" 
				Label100P.Text= i & "%"
			else
				Label = "Label" & i & "P"
				model.GetControlFromID(Label).BackColor="Silver" 
			end if		
		Next
	End If

	'Set required and disabled Controls

	ControlsCA = hlObj.GetContentIDs("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",0)
	For Each Control In ControlsCA

		If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0) = "ControlSettingRequired" Then
			
			If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeAttribute" Then
				model.GetControlFromID(hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0)).Font = LabelSubject.Font
				model.GetControlFromID(hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).Required = True
			End If

			If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeSearch" and hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0) <> "" Then
				model.GetControlFromID(hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0)).Font = LabelSubject.Font
			End If

		Else
			If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0) = "ControlSettingReadOnly" Then
				model.GetControlFromID(hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).Disabled = True
			End If
		End If

	Next

	'Button and Label switch in Classic and Web Desk
	Items = hlObj.GetItems(0,0,0,"StandardChangeCategory2StandardChange")

	If model.IsInWeb = false then

		LabelDocumentationSC.ShowControl = 3
		LabelCats.ShowControl = 3
		For Each Category In Items
			If Category.GetValue("StandardChangeCategorySpecific.DocumentationLink",0,0,0,0) = "" Then
				ButtonDocumentationSC.Disabled = true	
			Else
				ButtonDocumentationSC.Disabled = false	
			End If
		Next

	Else

		LabelCats.ShowControl = 1
		ButtonCats.Caption = ""
		For Each Category In Items
			If Category.GetValue("StandardChangeCategorySpecific.DocumentationLink",0,0,0,0) = "" Then
				LabelDocumentationSC.ShowControl = 3
				ButtonDocumentationSC.Caption = "Documentation"
				ButtonDocumentationSC.Disabled = true
			Else
				ButtonDocumentationSC.Caption = ""
				LabelDocumentationSC.ShowControl = 1
				LabelDocumentationSC.Text = "<a href='" + Category.GetValue("StandardChangeCategorySpecific.DocumentationLink",0,0,0,0) +"' target= '_blank' style = 'color: black'>Documentation</a>"
			End If
		Next

	End If

	'Helper Tabs für Gruppen 163392 und 163395 einblenden

	AgentID = hlSession.GetAgentID()
	Set Person = GetPersonForAgent(AgentID)

	GroupIDs = Person.GetValue("PersonDisplayHelper.GroupIDText",0,0,0,0)

	If inStr(GroupIDs, "163392") Or inStr(GroupIDs, "163395") Then
		TabPageAgent.ShowControl = 1
		TabPageReqAndDis.ShowControl = 1
	End If
End Function

Function StandardChangeOnSave()
	Dim count
	Dim varSubject
	Dim ControlsCA
	Dim go : go = true
	Dim Message : Message = ""

	varSubject = Left (TextBoxDescription.Text, 100)
	If TextBoxSubject.Text="" Then TextBoxSubject.Text = replace(varSubject,Chr(13)&Chr(10)," ")

	If IsObject(hlOrgunit) = true Then 
		hlObj.SetValue "RoutingHelper.GroupName",0,0,0,TextBoxGroupName.Text
		hlObj.SetValue "RoutingHelper.GroupID",0,0,0,hlOrgunit.ObjID
	End If


	If hlObj.IsNew = 1 then
		ButtonCloseTask.Disabled = true
		ButtonWaitingFor.Disabled = true
	End If


	'Check requiredControls

	ControlsCA = hlObj.GetContentIDs("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",0)
	For Each Control In ControlsCA

		If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0) = "ControlSettingRequired" Then
			
			If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeRelation" Then
				If hlObj.GetItemCount(&H00000, hlobj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,Control,0,0)) = 0 and hlObj.GetItemCount(&H10000, hlobj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,Control,0,0)) = 0 Then 
					go = false
					Message = Message & "Relation not set: " & model.GetControlFromID(hlobj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).Text & vbNewLine
				End If
			End If

			If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeSearch" and hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0) = "" Then
				If model.GetControlFromID(hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).GetSearchState <> 3 Then 
					go = False	
					Message = Message & "Search not set: " & model.GetControlFromID(hlobj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).SearchName & vbNewLine
				End If
			End If

		End If

	Next

	'Check Email Address of Additional Recipients
	If hlObj.GetItemCount(&H00000,"AdditionalRecipient2Case") > 0 Then
		For Each addrec in hlObj.GetItems(&H00000,0,0,"AdditionalRecipient2Case")
			If addrec.GetValue("PersonInformation.Email",0,0,0,0) = "" Then 
				hlObj.RemoveItem 0, addrec, "AdditionalRecipient2Case"
				Message = Message & "You selected an additional recipient that has no email address set. This additional recipient will be removed"
			End If
		Next
	End If

	If Message <> "" Then
		If go = False Then
			MsgBox Message 
			Model.CurrentCommand.Abort ""
		Else
			MsgBox Message 
		End If
	End If
End Function

Function StandardChangeCloseTask()
Dim task
Dim reservedBy
Dim taskStatus
Dim decision
Dim FlagManualDecision
Dim Go
Dim AgentID
Dim Person
Dim Groups
Dim Group
Dim count
Dim Agent
Dim ControlsCASTD 
Dim isnew : isnew = true

If model.Save() then
	If Not IsEmpty(TableControlTasks.SelectedObject) Then
			Set task = TableControlTasks.SelectedObject
			reservedBy = task.GetValue("CASEINFO.RESERVEDBY",0,0,0,0) 
			taskStatus = task.GetValue("CASEINFO.INTERNALSTATE",0,0,0,0) 
			decision = hlObj.GetValue("StandardChangeRecordSpecific.ManualDecision",0,0,0,0)
			FlagManualDecision = task.GetValue("TaskRecordSpecific.FlagManualDecision",0,0,0,0)
			AgentID = hlSession.GetAgentID()
			set Person = GetPersonForAgent(AgentID)


			If reservedBy <> "" Then
				model.MsgBox "Task is reserved by someone else"
			End If

			If taskStatus = "CLOSED" Then
				model.MsgBox "Task is already closed"
			End If

			If FlagManualDecision = 1 Then
				If decision = "" Then 
					model.MsgBox "Please select Answer first"
					Go = False
				Else
					Go = True
				End If
			Else
				Go = True
			End If

			If Go = True Then
				If taskStatus <> "CLOSED" AND ReservedBy = "" THEN
					task.Reserve
					Dim dtUntil, tmp
					tmp = 0
					dtUntil = DateAdd("s",3, Now)
					Do While DateDiff("s", Now, dtUntil) > 0
						tmp = tmp + 1
					Loop
					count = task.GetItemCount(&H00000,"Agent2Case")
					If count <> 0 then
						For i = 1 To count
							Agent = task.GetItems (0, -1, -1, "Agent2Case")
							task.RemoveItem 0, Agent(0), "Agent2Case"
						Next			
						task.AddItem 0,Person,"Agent2Case"
					Else
						task.AddItem 0,Person,"Agent2Case"
					End If
					
					task.SetValue "RoutingHelper.AgentName",0,0,0,Person.GetValue("PersonInformation.Name",0,0,task.GetSvcUnitCount(),0)
					task.SetValue "RoutingHelper.AgentID",0,0,0,Person.GetValue("HLOBJECTINFO.ID",0,0,GetSvcUnitCount(),0)
					task.SetValue "CaseSearchHelper.SolvedByAssignee",0,0,0,Person.GetValue("PersonInformation.Name",0,0,0,0)
					task.SetValue "CASEINFO.INTERNALSTATE",0,0,0,"CLOSED"
					task.SetValue "TaskRecordSpecific.Status",0,0,0,"StatusClosed"
					task.SetValue "TaskRecordSpecific.ManualDecision",0,0,0, decision
					model.SaveObject(task)
					task.Unreserve
				End If


				hlObj.SetValue "StandardChangeRecordSpecific.ManualDecision",0,0,0,""

				'Set required and disabled Controls

				ControlsCA = task.GetContentIDs("TaskRecordSpecific.RequiredAndDisabledControls_CA",0)
				For Each Control In ControlsCA

					If task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) <> "DataModelTypeAttribute" Then

						ControlsCASTD = hlObj.GetContentIDs("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",0)
						For Each ControlSTD In ControlsCASTD 
							If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,ControlSTD ,0,0) = task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,Control,0,0) Then
								isnew = False
								hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,ControlSTD,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0)
							End If	
						Next

						If isnew = True Then

							newID = hlObj.GenerateContentID()
							hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,newID,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) 
							hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,newID,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,Control,0,0) 
							hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,newID,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)
							hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,newID,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0)
							hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,newID,0, task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0)
	
						End If 

					End If

					If task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,0) = "ControlSettingRequired" Then
						
						If task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeAttribute" Then
							model.GetControlFromID(task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0)).Font = LabelSubject.Font
							model.GetControlFromID(task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).Required = True
						End If

						If task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,Control,0,0) = "DataModelTypeSearch" and task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0) <> "" Then
							model.GetControlFromID(task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,Control,0,0)).Font = LabelSubject.Font
						End If

					Else

						model.GetControlFromID(task.GetValue("TaskRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0)).Disabled = True

					End If

				Next

			End If
	End If
End If
End Function

Function StandardChangeWaitingFor()
	Dim task
	Set task = TableControlTasks.SelectedObject
	Dim reservedBy
	Dim taskStatus
	Dim decision
	Dim AgentID
	Dim Person

	reservedBy = task.GetValue("CASEINFO.RESERVEDBY",0,0,0,0) 
	taskStatus = task.GetValue("CASEINFO.INTERNALSTATE",0,0,0,0) 
	decision = hlObj.GetValue("StandardChangeRecordSpecific.ManualDecision",0,0,0,0)
	AgentID = hlSession.GetAgentID()
	set Person = GetPersonForAgent(AgentID)

	If reservedBy <> "" Then
		model.MsgBox "Task is reserved by someone else"
	End If

	If taskStatus = "CLOSED" Then
		model.MsgBox "Task is already closed"
	End If


	If taskStatus <> "CLOSED" AND ReservedBy = "" THEN
		task.Reserve

		Dim dtUntil, tmp
		tmp = 0
		dtUntil = DateAdd("s",3, Now)
		Do While DateDiff("s", Now, dtUntil) > 0
			tmp = tmp + 1
		Loop

		task.SetValue "CASEINFO.INTERNALSTATE",0,0,0,"WAITING_FOR"
		task.SetValue "TaskRecordSpecific.Status",0,0,0, "StatusWaitingFor"
		task.SetValue "TaskRecordSpecific.ManualDecision",0,0,0, decision
		count = task.GetItemCount(&H00000,"Agent2Case")
		If count <> 0 then
			For i = 1 To count
				Agent = task.GetItems (0, -1, -1, "Agent2Case")
				task.RemoveItem 0, Agent(0), "Agent2Case"
			Next			
			task.AddItem 0,Person,"Agent2Case"
		Else
			task.AddItem 0,Person,"Agent2Case"
		End If

		task.SetValue "RoutingHelper.AgentName",0,0,0,Person.GetValue("PersonInformation.Name",0,0,task.GetSvcUnitCount,0)
		task.SetValue "RoutingHelper.AgentID",0,0,0,Person.GetValue("HLOBJECTINFO.ID",0,0,GetSvcUnitCount,0)



		model.SaveObject(task)
		task.Unreserve
	End If

End Function

Sub RemoveRequiredOrDisabledField(controlname)

	Dim ControlsCA
	Dim removeID

	ControlsCA = hlObj.GetContentIDs("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",0)
	For Each Control In ControlsCA
		If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0) = controlname Then 
			removeID = Control
			Exit For
		End If 
	Next

	If removeID <> "" Then RemoveContentID "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",removeID,0 

End Sub



Sub AddRequiredOrDisabledField(datamodeltype, datamodelname, controlname, controllabelname, controlsetting)

	Dim ControlsCA
	Dim newID
	Dim isnew : isnew = true

	ControlsCA = hlObj.GetContentIDs("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA",0)
	For Each Control In ControlsCA
		If hlObj.GetValue("StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,Control,0,0) = controlname Then
			isnew = False
			hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,Control,0,controlsetting
			Exit For
		End If	
	Next 

	If isnew = True Then

		newID = hlObj.GenerateContentID()
		hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelType",0,newID,0, datamodeltype
		hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.DataModelName",0,newID,0, datamodelname
		hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.Name",0,newID,0, controlname
		hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.LabelName",0,newID,0, controllabelname
		hlObj.SetValue "StandardChangeRecordSpecific.RequiredAndDisabledControls_CA.ControlSetting",0,newID,0, controlsetting

	End If 

End Sub
