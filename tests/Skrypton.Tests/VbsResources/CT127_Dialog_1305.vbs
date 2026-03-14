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