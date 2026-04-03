ButtonOutlookExport_Click
SUB ButtonOutlookExport_Click()	
	Dim olApp 'As New Outlook.Application 
	Dim olContact 'As Outlook.ContactItem 
	Dim olNS 'As Outlook.NameSpace 
	Dim olFolder 'As Outlook.MAPIFolder 
	
	Dim strFirstName 'As String 
	Dim strLastName 'As String
	Dim AddressStreet 
	Dim AddressCountry 
	Dim AddressPostalCode
	Dim AddressCity 
	Dim TelephoneNumber 
	Dim Email1Address 
	Dim MobileTelephoneNumber
	Dim CompanyName 
 
	Dim blnContinue 'As Boolean 
	Dim varReturnVal 'As Variant 
	Dim folderint
	
	
	Set olApp = CreateObject("Outlook.Application") 
		
	folderint = 10
	
	Set olContact = olApp.CreateItem(olContactItem) 
	Set olNS = olApp.GetNamespace("MAPI") 
	Set olFolder = olNS.GetDefaultFolder(folderint) 
	
	blnContinue = True 
	
	' Suche ob der Kontakt schon existiert
	
		'Set olContact = olFolder.Items.Find("[lastname] = " & strLastName) 
		Set olContact = olFolder.Items.Find("[lastname] = '" & strLastName & "'  And [firstname] = '" & strFirstName & "'") 
		If Not TypeName(olContact) = "Nothing" Then  'Match found 
			
			varReturnVal = MsgBox("Es exsitiert bereits ein Datensatz mit dem gleichen Namen. soll der Kontakt dennoch angelegt werden?", vbOKCancel, "Doppelter Eintrag?") 
			
			If varReturnVal = vbCancel Then 
				blnContinue = False 
			End If 
		End If 	
	
	' Kontakt anlegen falls notwendig
	If blnContinue = True Then
		'On Error GoTo Error_Handler 
		Const olContactItem = 2 
	
		Set olContact = olApp.CreateItem(olContactItem)   
		With olContact 
			.FirstName = strFirstName
			.LastName = strLastName
			'.JobTitle = "" 
			.CompanyName = CompanyName
			.BusinessAddressStreet = AddressStreet
			.BusinessAddressCity = AddressCity
			'.BusinessAddressState = "Quebec" 
			.BusinessAddressCountry = AddressCountry
			.BusinessAddressPostalCode = AddressPostalCode
			.BusinessTelephoneNumber = TelephoneNumber
			.BusinessFaxNumber = "" 
			.Email1Address = Email1Address
			.MobileTelephoneNumber = MobileTelephoneNumber 
			.Save 'use 
			.Display 'if you wish the user to see the contact pop-up 
		End With   
	End If	
	
END SUB