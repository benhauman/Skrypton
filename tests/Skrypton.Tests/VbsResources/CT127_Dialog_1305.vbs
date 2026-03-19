SUB ComboBoxAccountedPermissionType_ondatachange()
END SUB
SUB OnUpdate()
END SUB
SUB ComboBoxAccountedPermissionType_SelectionEndOK()
END SUB
SUB ComboBoxAccountedPermissionType_SelectionChanged()
	If ComboBoxAccountedPermissionType.Text = "LDAP" Then
	
	TabPageLdap.ShowControl = 1
	TabPageCsv.ShowControl = 3
	TextBoxUsed.ShowControl = 1
	LabelUsed.ShowControl = 1
	TableControlCostCenter.ShowControl = 1
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "CSV" Then
	
	TabPageCsv.ShowControl = 1
	TabPageLdap.ShowControl = 3
	TextBoxUsed.ShowControl = 1
	LabelUsed.ShowControl = 1
	TableControlCostCenter.ShowControl = 3
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "Other" Then
	
	TabPageCsv.ShowControl = 3
	TabPageLdap.ShowControl = 3
	TextBoxUsed.ShowControl = 3
	LabelUsed.ShowControl = 3
	TableControlCostCenter.ShowControl = 3
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "SAP" Then
	
	TabPageCsv.ShowControl = 3
	TabPageLdap.ShowControl = 3
	TextBoxUsed.ShowControl = 3
	LabelUsed.ShowControl = 3
	TableControlCostCenter.ShowControl = 3
	
	End If 

END SUB
SUB ComboBoxAccountedPermissionType_onfocus()
END SUB
SUB OnLoad()
	If ComboBoxAccountedPermissionType.Text = "LDAP" Then
	
	TabPageLdap.ShowControl = 1
	TextBoxUsed.ShowControl = 1
	LabelUsed.ShowControl = 1
	TabPageCsv.ShowControl = 3
	TableControlCostCenter.ShowControl = 1
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "CSV" Then
	
	TabPageCsv.ShowControl = 1
	TextBoxUsed.ShowControl = 1
	LabelUsed.ShowControl = 1
	TabPageLdap.ShowControl = 3
	TableControlCostCenter.ShowControl = 3
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "Other" Then
	
	TabPageCsv.ShowControl = 3
	TabPageLdap.ShowControl = 3
	TextBoxUsed.ShowControl = 3
	LabelUsed.ShowControl = 3
	TableControlCostCenter.ShowControl = 3
	
	End If 
	
	If ComboBoxAccountedPermissionType.Text = "SAP" Then
	
	TabPageCsv.ShowControl = 3
	TabPageLdap.ShowControl = 3
	TextBoxUsed.ShowControl = 3
	LabelUsed.ShowControl = 3
	TableControlCostCenter.ShowControl = 3
	
	End If 

END SUB
SUB OnSave()
	Dim TypeLib 
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	
	Dim Guid
	Guid = hlObj.GetValue ("ExternalReference.ExternalRefNo",0,0,0,0) 
	
	If (Guid = "") Then
	hlObj.SetValue "ExternalReference.ExternalRefNo", 0, 0, 0, Left(TypeLib.Guid, 38)
	End If
	
END SUB
