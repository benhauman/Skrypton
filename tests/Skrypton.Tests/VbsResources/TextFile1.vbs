'---------------------------------------------------------------
'Diese Funktion ermittelt den Standard-Eintrag zum angegebenen Attribut aus
'dem Dictionary.
'Wenn der Parameter "GetAll" auf False steht wird als Rückgabewert für die Funktion
'ebenfalls "False" ausgegben, wenn mehr als ein Standardeintrag gefunden wird.
'Wenn für den Parameter "True" angeben wird, prüft die Funktion ob es tatsächlich
'nur einen Standard-Eintrag gibt, sonst "False".
Function GetCommunicationDefault(hlContext, hlObject, dict, GetAll)
        GetCommunicationDefault = False
        Dim ItemCount: ItemCount = 0
        Dim strValue: strValue = ""
        
        Dim ItemIDs: ItemIDs = ""
        ItemIDs = hlObject.GetContentIDs(dict("Compound"), 0)

        Dim Item: Item = 0
        For Each Item In ItemIDs
                Dim defItem: defItem = False
                defItem = GetFlagValue(hlContext, hlObject, dict("Default"), Item, 0)
                If CBool(defItem) = True Then 
                        ItemCount = ItemCount + 1
                        strValue = hlObject.GetValue(dict("Value"), 0, Item, 0, 0)
                        If CBool(GetAll) = False Then Exit For
                End If
        Next
        If ItemCount > 1 Then
                GetCommunicationDefault = False
                Exit Function
        Else
                GetCommunicationDefault = True
                dict("DefValue") = strValue
        End If
End Function
'----------------------------------------------------------------------------------------------------------
'Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
Sub Trace(hlContext,text)
  hlContext.trace 1, text
End Sub
'---------------------------------------------------------------
'Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
Sub SetPersonInformation(hlContext, hlObject, dict)
        'Aus dem Dictionary wird das Attribut und der dazugehörige Wert ermittelt.
        Dim AttrDef: AttrDef = ""
        AttrDef = "PersonInformation." & dict("PersInfoAttr")
        
        Dim strAttrValue: strAttrValue = ""
        strAttrValue = dict("DefValue")
        
	If strAttrValue = "" Then strAttrValue = "-"
	hlObject.SetValue AttrDef, 0, 0, 0, strAttrValue
End Sub
'---------------------------------------------------------------
Function IsHLObject(hlContext, hlObject)
'	Trace hlContext, "IsObject " & IsObject(hlObject)
'	Trace hlContext, "IsNull " & IsNull(hlObject) 
'	Trace hlContext, "IsEmpty " & IsEmpty(hlObject)
'	Trace hlContext, "Leerstring " 
'	Trace hlContext, "Leerstring " & hlObject = ""
	Trace hlContext, "Type " & VarType(hlObject)
	IsHLObject = ((IsObject(hlObject) = True) And ((hlObject Is Nothing) = False))
End Function
'-------------------------------------------------------------------
Function GetBaseType(hlContext, hlObject)
        GetBaseType = hlObject.GetValue("HLOBJECTINFO.BASETYPE", 0, 0, 0, 0)
End Function
'---------------------------------------------------------------
'Dies ist eine rekursive Function zum ermitteln der Organisationshierarchie,
'ausgehend vom der ersten OU überhalb einer Person.
'Die Variable "strOrgUnits" ist der Out-Parameter der Function.
Function GetPersonOrganisation(hlContext, hlOrgUnit, strOrgUnits)
        GetPersonOrganisation = 0
        Dim retval: retval = 0

        'Wenn noch keine OU ermittelt wurde, wird der Name der ersten OU eingetragen.
        'Andernfalls, wird jede weitere OU einfach angehangen.
        If strOrgUnits = "" Then
                strOrgUnits = hlOrgUnit.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
        Else
                strOrgUnits = strOrgUnits & ", " & hlOrgUnit.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
        End If

        'Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
        'für die nächste Abfrage gewählt werden kann.
	Dim NextOrgUnit
      Dim orgaType : orgaType = ""
      orgaType = hlOrgUnit.GetType
      If orgaType = "Division" Then
		NextOrgUnit = hlOrgUnit.GetItems(&H10000, 0, 0, "CompanyView")
	End If
	If orgaType = "Site" then
		NextOrgUnit = hlOrgUnit.GetItems(&H10000, 0, 0, "Site2Company")
	End If
	If orgaType = "Company" Then
		NextOrgUnit = hlOrgUnit.GetItems(&H10000, 0, 0, "Company2Company")
	End If

        'Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
        'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
	If IsArray(NextOrgUnit) Then
	      If UBound(NextOrgUnit) >= 0 Then
      	      retval = GetPersonOrganisation(hlContext, NextOrgUnit(0), strOrgUnits)
	      Else
            	Exit Function
      	End If
	End If
End Function
'---------------------------------------------------------------
'Über diese Function wird für ein Flag Attribut immer der Wert
'True oder False ausgegeben.
Function GetFlagValue(hlContext, hlObject, hlattribute, hlcontentid, hlsuid)
        GetFlagValue = hlObject.GetValue(hlattribute, 0, hlcontentid, hlsuid, 0)
        If GetFlagValue = "" Then GetFlagValue = False
End Function
'-------------------------------------------------------------------
'Diese Function ermitellt eine Fehlermeldung aus dem helpLine
'Wörterbuch ohne Parameter.
Function GetErrMsg0(hlContext, LocaleID, ErrCode)
        GetErrMsg0 = ""

        Dim strErrMsg: strErrMsg = ""
        strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
        strErrMsg = strErrMsg & vbNewLine & "(Code: " & ErrCode & ")"

        'Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
        'Rückgabewert der Function ist die Fehlermeldung.
        GetErrMsg0 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
'---------------------------------------------------------------

'Das Script ermittelt auf Basis der ersten übergeordneten OU den gesamten Pfad bis zur Firma oder Konzern
'und speichert diesen in das Hilfsattribut PersonInformation.PersonOrganisation.
'This script detects the entire path based on the first parent OU up to the company or holding
'and saves them into the attribute PersonInformation.PersonOrganisation.
Sub SetPersonOrganization(hlContext, hlPerson, dict)
	Dim FirstOrgUnit : Set FirstOrgUnit = Nothing
	Set FirstOrgUnit = hlContext.GetRelatedObject

	if IsHLObject(hlContext, FirstOrgUnit) = True Then
		If FirstOrgUnit.GetType <> "Company" And FirstOrgUnit.GetType <> "Division" Then
			Set FirstOrgUnit = Nothing
		End If
	End if

	If IsHLObject(hlContext, FirstOrgUnit) = False Then
		Dim rsltOrgUnit : rsltOrgUnit = ""
		rsltOrgUnit = hlPerson.GetItems(&H10000, 0, 0, "Person2Organization")
		If UBound(rsltOrgUnit) >=0 Then
			Set FirstOrgUnit = rsltOrgUnit(0)
		End If
	End If

	If IsHLObject(hlContext, FirstOrgUnit) = True Then
		If GetBaseType(hlContext, FirstOrgUnit) = "ORGANISATION" Then
			Dim retval : retval = ""
			Dim strOrgUnits : strOrgUnits = ""
			retval = GetPersonOrganisation(hlContext, FirstOrgUnit, strOrgUnits)

			dict("DefValue") = strOrgUnits
			dict("PersInfoAttr") = "PersonOrganization"
			Call SetPersonInformation(hlContext, hlPerson, dict)
		End If
	End If
End Sub
'---------------------------------------------------------------
'SACM
'----------------------------------------------------------------------------------------------------------
'Globale Konstanten für freie Assoziationsdefinitionen
Const HLASC_SoftwareLicenseFolderView = "LicenseFolderView"
Const HLASC_SoftwareLicenseGroupView = "LicenseGroupView"
Const HLASC_Software2Computer = "Software2Computer"
'----------------------------------------------------------------------------------------------------------
'Prozedur füllt die Umzugshistorie für das entsprechende Objekt
Sub SetAssetHistory(hlContext, hlObjectA, hlObjectB, created)
    
    Dim productDefName
    productDefName = hlObjectB.GetType()
    
    If (productDefName <> "Software" And productDefName <> "SoftwareLicence") Then
        Dim agentID, contentID, personOfAgent, personName, orgUnitName
        contentID = hlObjectB.GenerateContentID()
        agentID = hlContext.GetAgentID()
        orgUnitName = hlObjectA.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
        Set personOfAgent = hlContext.GetPersonOfAgent(agentID)
        If (personOfAgent Is Nothing) Then
            Dim strErrMsg
            strErrMsg = GetErrMsg0(hlContext, hlContext.GetLocaleID, "#ERR_SETASSETHISTORY")
            Trace hlContext, strErrMsg
            'hlContext.abortcommand strErrMsg
        Else
            personName = personOfAgent.GetValue("PersonGeneral.Name", 0, 0, 0, 0)
            personName = personName & ", "
            personName = personName & personOfAgent.GetValue("PersonGeneral.GivenName", 0, 0, 0, 0)
        End If
        hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangedBy", 0, contentID, 0, personName
        hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID", 0, contentID, 0, agentID
        hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangeDate", 0, contentID, 0, Now()
        hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryOrgUnit", 0, contentID, 0, orgUnitName
        hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryOrgUnitID", 0, contentID, 0, hlObjectA.GetID()
        
        If (created = True) Then
            hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryAction", 0, contentID, 0, "HistoryActionCreated"
        Else
            hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryAction", 0, contentID, 0, "HistoryActionDeleted"
        End If
    End If
End Sub
'---------------------------------------------------------------
'Diese Function ermitellt eine Fehlermeldung aus dem helpLine
'Wörterbuch mit einem Parameter.
Function GetErrMsg1(hlContext, LocaleID, ErrCode, Arg1)
        GetErrMsg1 = ""

        Dim strErrMsg: strErrMsg = ""
        strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
        strErrMsg = Replace(strErrMsg, "%1", Arg1)
        strErrMsg = strErrMsg & vbLf & "(Code: " & ErrCode & ")"
        
        'Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
        'Rückgabewert der Function ist die Fehlermeldung.
        GetErrMsg1 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
Function GetErrMsg2(hlContext, LocaleID, ErrCode, Arg1, Arg2)
        GetErrMsg2 = ""

        Dim strErrMsg: strErrMsg = ""
        strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
        strErrMsg = Replace(strErrMsg, "%1", Arg1)
        strErrMsg = Replace(strErrMsg, "%2", Arg2)
        strErrMsg = strErrMsg & vbLf & "(Code: " & ErrCode & ")"
        
        'Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
        'Rückgabewert der Function ist die Fehlermeldung.
        GetErrMsg2 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Funktion wird geprüft, ob es unterhalb einer Software Suite
'bereits Lizenzumschläge mit Lizenzen gibt.
Function GetReferenceLicenseCount(hlContext, hlSWFolder, chkFolderOnly, HLASC_SoftwareLicenseFolderView)
        GetReferenceLicenseCount = 0
        
        Dim rsltSWFolders: rsltSWFolders = ""
        Dim SoftwareLicense: Set SoftwareLicense = Nothing
        Dim objType: objType = ""

        'Prüfen ob es Software Lizenzobjekte/Lizenzumschläge unterhalb des Folders gibt.
        rsltSWFolders = hlSWFolder.GetItems(&H0, -1, -1, HLASC_SoftwareLicenseFolderView)

        For Each SoftwareLicense In rsltSWFolders
                objType = SoftwareLicense.GetType()
                If objType = "LicenseFolder" Then
                        GetReferenceLicenseCount = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                        If GetReferenceLicenseCount > 0 Then
                                Exit Function
                        End If
                End If
                If objType = "SoftwareLicense" And CBool(chkFolderOnly) = False Then
                        GetReferenceLicenseCount = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                        If GetReferenceLicenseCount > 0 Then
                                Exit Function
                        End If
                End If
        Next
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
'den obersten Lizenz Umschlag ermittelt. Auf dem Weg dort hin wird geprüft ob einer
'der Lizenzumschläge eine Software Suite ist.
Function CheckForSoftwareSuiteFolder(hlContext, hlParentSWFolder, pDict, HLASC_SoftwareLicenseFolderView)
    CheckForSoftwareSuiteFolder = ""
    Dim retval : retval = 0
    Dim NextSWFolder : NextSWFolder = ""

    'Festhalten auf welcher Ebene ggf. eine Software Suite oberhalb des
    'Start Folders existiert. Die Variable muss von außen mit einem Startwert
    'initialisiert werden.
    If pDict("SoftwareSuiteFolderLevel") = 0 Or pDict("SoftwareSuiteFolderLevel") = "" Then
        pDict("SoftwareSuiteFolderLevel") = 1
    Else
        pDict("SoftwareSuiteFolderLevel") = pDict("SoftwareSuiteFolderLevel") + 1
    End If

    'Amhand des Flags "Software Suite" festellen ob ein Lizenzumschlag als Software Suite
    'gekennzeichnet ist. Falls Ja, Name des Umschlags auslesen und Funktion abbrechen.
    Dim CheckSoftwareSuite : CheckSoftwareSuite = False
    CheckSoftwareSuite = GetFlagValue(hlContext, hlParentSWFolder, "SoftwareLicenseFolderDetail.FlagSoftwareSuite", 0, 0)
    If CBool(CheckSoftwareSuite) = True Then
        pDict("SoftwareSuiteFolder") = hlParentSWFolder.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
        Exit Function
    End If

    'Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
    'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
    NextSWFolder = hlParentSWFolder.GetItems(&H10000, -1, -1, HLASC_SoftwareLicenseFolderView)
    If UBound(NextSWFolder) >= 0 Then
        retval = CheckForSoftwareSuiteFolder(hlContext, NextSWFolder(0), pDict, HLASC_SoftwareLicenseFolderView)
    Else
        Exit Function
    End If
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
'den obersten Lizenz Umschlag ermittelt und neu berechnet hat.
Function SetLicenseCounter(hlContext, hlSWFolder, pDict, assocName)
        SetLicenseCounter = 0
        Dim retval: retval = 0
        
        'Dictionary Einträge initalisieren
        pDict("SoftwareLicenses") = ""
        pDict("SumRefLicCounter") = 0
        pDict("SumInstLicCounter") = 0
        pDict("SumFreeLicCounter") = 0
        
        'Prüfen ob es Software Lizenzobjekte unterhalb des Folders gibt.
        pDict("SoftwareLicenses") = hlSWFolder.GetItems(&H0, -1, -1, assocName)
        
        'Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
        'Objekte gezählt werden müssen
        Dim CheckSoftwareSuite: CheckSoftwareSuite = False
        CheckSoftwareSuite = GetFlagValue(hlContext, hlSWFolder, "SoftwareLicenseFolderDetail.FlagSoftwareSuite", 0, 0)
        
        If UBound(pDict("SoftwareLicenses")) >= 0 Then
                If CBool(CheckSoftwareSuite) = False Then
                        Call CalcAllLicCounter(hlContext, pDict)
                Else
                        Call CalcFolderLicCounter(hlContext, pDict)
                End If
        End If
        'Gesatmzahl der Lizenzen in den Lizenzumschlag zurückschreiben
        hlSWFolder.SetValue "SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, pDict("SumRefLicCounter")
        hlSWFolder.SetValue "SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, pDict("SumInstLicCounter")
        
        'Wenn die Lizenzkontrolle durch den Applikations Server erfolgt ("Lizenzkontrolle durch Server")
        'dann die Anzahl freier Lizenzen immer auf den Wert "0" setzen.
        Dim CheckLicContrByServer: CheckLicContrByServer = False
        CheckLicContrByServer = GetFlagValue(hlContext, hlSWFolder, "SoftwareLicenseFolderDetail.FlagLicenseControlledByServer", 0, 0)
        If CBool(CheckLicContrByServer) = True Then
                pDict("SumFreeLicCounter") = 0
        End If
        hlSWFolder.SetValue "SoftwareLicenseCounter.FreeLicenseCount", 0, 0, 0, pDict("SumFreeLicCounter")

        'Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
        'für die nächste Abfrage gewählt werden kann.
        Dim NextSWFolder: NextSWFolder = ""
        Dim a: a = ""
        a = hlSWFolder.GetType
        If a = "LicenseFolder" Then
                NextSWFolder = hlSWFolder.GetItems(&H10000, 0, 0, assocName)
        End If
        'Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
        'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
        If UBound(NextSWFolder) >= 0 Then
                retval = SetLicenseCounter(hlContext, NextSWFolder(0), pDict, assocName)
        Else
                Exit Function
        End If
End Function
'----------------------------------------------------------------------------------------------------------
Function IsValidObject(obj)
 IsValidObject = (IsObject(obj) And ( Not (obj Is Nothing)) )
End Function
'----------------------------------------------------------------------------------------------------------
Sub CalcAllLicCounter(hlContext, pDict)
        Dim SWRefLicCounter: SWRefLicCounter = 0
        Dim SWInstCounter: SWInstCounter = 0
        Dim SoftwareLicense: Set SoftwareLicense = Nothing
        Dim objType: objType = ""
        Dim lstLicStatus: lstLicStatus = ""
        
        For Each SoftwareLicense In pDict("SoftwareLicenses")
                objType = SoftwareLicense.GetType()
                If objType = "SoftwareLicense" Then
                        lstLicStatus = SoftwareLicense.GetValue("SoftwareLicenseDetail.LicenseStatus", 0, 0, 0, 0)
                        If lstLicStatus = "LicenseStatusValid" Then
                                SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                                pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
                        End If
                Else
                        If objType = "LicenseFolder" Or objType = "Software" Then
                                SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                                pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
                                SWInstCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, 0))
                                pDict("SumInstLicCounter") = pDict("SumInstLicCounter") + SWInstCounter
                        End If
                End If
        Next
        'Anzahl freier Lizenzen berechnen und in den Folder schreiben.
        pDict("SumFreeLicCounter") = pDict("SumRefLicCounter") - pDict("SumInstLicCounter")

End Sub
'----------------------------------------------------------------------------------------------------------
Sub CalcFolderLicCounter(hlContext, pDict)

        Dim SWRefLicCounter: SWRefLicCounter = 0
        Dim SWInstCounter: SWInstCounter = 0
        Dim SoftwareLicense: Set SoftwareLicense = Nothing
        Dim objType: objType = ""
        Dim lstLicStatus: lstLicStatus = ""
        
        For Each SoftwareLicense In pDict("SoftwareLicenses")
                objType = SoftwareLicense.GetType()
                If objType = "LicenseFolder" Or objType = "Software" Then
                        SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                        pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter

                        SWInstCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, 0))
                        If SWInstCounter > pDict("SumInstLicCounter") Then
                                pDict("SumInstLicCounter") = SWInstCounter
                        End If
                End If
                If objType = "SoftwareLicense" Then
                        lstLicStatus = SoftwareLicense.GetValue("SoftwareLicenseDetail.LicenseStatus", 0, 0, 0, 0)
                        If lstLicStatus = "LicenseStatusValid" Then
                                SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
                                pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
                        End If
                End If
        Next
        'Anzahl freier Lizenzen berechnen und in den Folder schreiben.
        pDict("SumFreeLicCounter") = pDict("SumRefLicCounter") - pDict("SumInstLicCounter")
End Sub
'----------------------------------------------------------------------------------------------------------
'Diese Function überprüft den ganzzahligen Wert (Integer).
Function CheckIntegerValue(hlContext, intval)
    If intval = "" Or IsNumeric(intval) = False Then
        CheckIntegerValue = 0
    Else
        CheckIntegerValue = CLng(intval)
    End If
End Function
'----------------------------------------------------------------------------------------------------------
Function OnCreate_HasAssociationToDelete(hlContext, AscDefName, hlObjB)
        Dim result: result = False
        Dim cAssociationChanges: cAssociationChanges = 0
        cAssociationChanges = hlContext.GetAssociationChangesCount

        Dim oAssociationChange: Set oAssociationChange = Nothing
        Dim AscDefNameChange: AscDefNameChange = ""
        Dim ixAC: ixAC = 0
        
        For ixAC = 0 To cAssociationChanges - 1
                Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)
          
                AscDefNameChange = oAssociationChange.AssociationType
                
                If oAssociationChange.IsToDelete Then
                        If (AscDefNameChange = AscDefName) Then
                                If (hlObjB.GetID = oAssociationChange.EndB.GetID) Then
                                        result = True
                                        Exit For
                                End If 'check the ids
                        End If ' check the defnames
                End If ' is to create
        Next
        OnCreate_HasAssociationToDelete = result
End Function
'----------------------------------------------------------------------------------------------------------
Function OnCreate_HasAssociationToCreate(hlContext, AscDefName, hlObjB)
        Dim result: result = False
        Dim cAssociationChanges: cAssociationChanges = 0
        cAssociationChanges = hlContext.GetAssociationChangesCount

        Dim oAssociationChange: Set oAssociationChange = Nothing
        Dim AscDefNameChange: AscDefNameChange = ""
        Dim ixAC: ixAC = 0
        
        For ixAC = 0 To cAssociationChanges - 1
                Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)
          
                AscDefNameChange = oAssociationChange.AssociationType
                
                If oAssociationChange.IsToCreate Then
                        If (AscDefNameChange = AscDefName) Then
                                If (hlObjB.GetID = oAssociationChange.EndB.GetID) Then
                                        result = True
                                        Exit For
                                End If 'check the ids
                        End If ' check the defnames
                End If ' is to create
        Next
        OnCreate_HasAssociationToCreate = result
End Function

Function OnDelete_HasAssociationToCreate(hlContext, AscDefName, hlObjB) ' bool
        Dim result: result = False
        
        'Anzahl der zu erstellenden oder löschenden Assoziationen
        Dim cAssociationChanges: cAssociationChanges = 0
        cAssociationChanges = hlContext.GetAssociationChangesCount

        Dim oAssociationChange: Set oAssociationChange = Nothing
        Dim AscDefNameChange: AscDefNameChange = ""
        Dim ixAC: ixAC = 0
        
        For ixAC = 0 To cAssociationChanges - 1

                'Für jede Assoziations Änderung wird das entsprechende Infos (Objekt    ) ausgelsen.
                Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)
                'Def Name der Assoc ermitteln, die angelegt werden soll
                AscDefNameChange = oAssociationChange.AssociationType

                If oAssociationChange.IsToCreate Then
                        'Überprüfen ob die gewünschte Assoc auch angelegt werden soll.
                        If (AscDefNameChange = AscDefName) Then
                                If (hlObjB.GetID = oAssociationChange.EndB.GetID) Then
                                        result = True
                                        Exit For
                                End If 'check the ids
                        End If ' check the defnames
                End If ' is to create
        Next
        OnDelete_HasAssociationToCreate = result
End Function
'----------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------
Function GetAssociatedOrganizationalUnit(hlContext, lcid, hlChild, pDict, outParentDefName)
        GetAssociatedOrganizationalUnit = ""
        outParentDefName = ""
        
        Dim rsltParent: rsltParent = ""
        rsltParent = hlChild.GetItems(&H10000, -1, -1, pDict("AssocID"))
        If UBound(rsltParent) >= pDict("ParentCounter") Then
                Dim objParent: Set objParent = Nothing
                For Each objParent In rsltParent
                        GetAssociatedOrganizationalUnit = objParent.GetValue(pDict("AttrName"),0, 0, 0, 0)
                        outParentDefName = hlContext.GetDisplayName(objParent.GetValue("HLOBJECTINFO.DEFID",0,0,0,0), lcid)
                        Exit For
                Next
        End If
End Function
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
Function MIG_CreateXMLDocument (hlSrvContext, pDict)
						
		'XML-Objekt erstellen
		Dim objXMLDoc : Set objXMLDoc = Nothing
		Set objXMLDoc = CreateObject("Msxml2.DOMDocument")
		
		'XML-Processing Instruction hinzufügen
		Dim xmlProInc : Set xmlProInc = Nothing
		Set xmlProInc = objXMLDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
		objXMLDoc.insertBefore xmlProInc, objXMLDoc.firstChild
		
		'Root-Element erstellen
		Dim xmlRoot : Set xmlRoot = objXMLDoc.CreateElement("ASAPBatch")
		objXMLDoc.AppendChild(xmlRoot)
		xmlRoot.SetAttribute "xmlns", "http://www.brainware.ch/operationsmanager/asap-batch/1.1"
		xmlRoot.SetAttribute "xmlns:dt", "http://www.brainware.ch/operationsmanager/wf/changemanagement/columbus/datatypes/1.1"
		xmlRoot.SetAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance" 
		xmlRoot.SetAttribute "xsi:schemaLocation", "http://www.brainware.ch/operationsmanager/asap-batch/1.1 asap-batch-1.1.xsd" 
		xmlRoot.SetAttribute "version", "1.1" 
		xmlRoot.SetAttribute "responseRequired", "Yes"
		
		'Das Node Session hinzufügen
		Dim nodeSession : Set nodeSession = objXMLDoc.CreateElement("Session")
		xmlRoot.AppendChild (nodeSession)
		nodeSession.SetAttribute "id", "s1"
		nodeSession.SetAttribute "loginname", "foreignSystems\assetcolumbus"
		nodeSession.SetAttribute "password", ""
		
		'XML Dokument inkl. Header an das Dictionary übergeben
		Set pDict("XMLDocument") = objXMLDoc
End Function
'---------------------------------------------------------------------
Function MIG_CreateADDXML2Columbus (hlSrvContext, pDict)
		
		'Root Element aus dem XML ermitteln.
		Dim xmlRoot : Set xmlRoot = pDict("XMLDocument").DocumentElement
		
		'Das Node CreateInstanceReq hinzufügen
		Dim nodeCreateInstanceRq : Set nodeCreateInstanceRq = pDict("XMLDocument").CreateElement("CreateInstanceRq")
		xmlRoot.AppendChild (nodeCreateInstanceRq)
		nodeCreateInstanceRq.SetAttribute "id", "e7"
		nodeCreateInstanceRq.SetAttribute "wfpNs", "ch.bw.wf.changemgmt.columbus_adddevice"
		nodeCreateInstanceRq.SetAttribute "wfmNs", "Columbus Changemanagement"
		nodeCreateInstanceRq.SetAttribute "sessionId", "s1"
		
		'Das Node ObserverKey hinzufügen
		Dim nodeObserverKey : Set nodeObserverKey = pDict("XMLDocument").CreateElement("ObserverKey")
		nodeCreateInstanceRq.AppendChild (nodeObserverKey)
		nodeObserverKey.Text = pDict("ObserverKey")

		'Das Container Node ContextData hinzufügen
		Dim nodeContextData : Set nodeContextData = pDict("XMLDocument").CreateElement("ContextData")
		nodeCreateInstanceRq.AppendChild (nodeContextData)
		
		'Das Container Node AddDeviceActualParams hinzufügen
		Dim nodeAddDeviceActualParams : Set nodeAddDeviceActualParams = pDict("XMLDocument").CreateElement("dt:AddDeviceActualParams")
		nodeContextData.AppendChild (nodeAddDeviceActualParams)
		
		'Das Container Node DeviceIdentification hinzufügen
		Dim nodeDeviceIdentification : Set nodeDeviceIdentification = pDict("XMLDocument").CreateElement("dt:DeviceIdentification")
		nodeAddDeviceActualParams.AppendChild (nodeDeviceIdentification)
		
		'Das Node DeviceName hinzufügen
		Dim nodeDeviceName : Set nodeDeviceName = pDict("XMLDocument").CreateElement("dt:DeviceName")
		nodeDeviceIdentification.AppendChild (nodeDeviceName)
		nodeDeviceName.Text = pDict("DeviceName")
		
		'Das Node CompanyName hinzufügen
		Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
		nodeDeviceIdentification.AppendChild (nodeCmpyName)
		nodeCmpyName.Text = pDict("CompanyName")
		
		'Das Node Domain hinzufügen
		Dim nodeDomain : Set nodeDomain = pDict("XMLDocument").CreateElement("dt:Domain")
		nodeDeviceIdentification.AppendChild (nodeDomain)
		nodeDomain.Text = pDict("Domain")
		
		'Das Node CostCenter hinzufügen
		Dim nodeCostCenter : Set nodeCostCenter = pDict("XMLDocument").CreateElement("dt:CostCenter")
		nodeAddDeviceActualParams.AppendChild (nodeCostCenter)
		nodeCostCenter.Text = pDict("CostCenter")
		
		'Das Node MACAdess hinzufügen
		Dim nodeMACAddress : Set nodeMACAddress = pDict("XMLDocument").CreateElement("dt:MACAddress")
		nodeAddDeviceActualParams.AppendChild (nodeMACAddress)
		nodeMACAddress.Text = pDict("MACAddress")
		
		'Das Node SubnetMask hinzufügen
		Dim nodeSubnetMask : Set nodeSubnetMask = pDict("XMLDocument").CreateElement("dt:SubnetMask")
		nodeAddDeviceActualParams.AppendChild (nodeSubnetMask)
		nodeSubnetMask.Text = pDict("SubnetMask")
		
		'Das Node HwTypeId hinzufügen
		Dim nodeHWType : Set nodeHWType = pDict("XMLDocument").CreateElement("dt:HwTypeId")
		nodeAddDeviceActualParams.AppendChild (nodeHWType)
		nodeHWType.Text = pDict("HwTypeId")
		
		'Das Node OsTypeId hinzufügen
		Dim nodeOSType : Set nodeOSType = pDict("XMLDocument").CreateElement("dt:OsTypeId")
		nodeAddDeviceActualParams.AppendChild (nodeOSType)
		nodeOSType.Text = pDict("OsTypeId")
		
		'Das Node ActivationState hinzufügen
		Dim nodeActState : Set nodeActState = pDict("XMLDocument").CreateElement("dt:ActivationState")
		nodeAddDeviceActualParams.AppendChild (nodeActState)
		nodeActState.Text = pDict("ActivationState")

End Function
'---------------------------------------------------------------------
Function MIG_CreateCHGXML2Columbus (hlSrvContext, pDict)
						
		'Root Element aus dem XML ermitteln.
		Dim xmlRoot : Set xmlRoot = pDict("XMLDocument").DocumentElement
		
		'Das Node CreateInstanceReq hinzufügen
		Dim nodeCreateInstanceRq : Set nodeCreateInstanceRq = pDict("XMLDocument").CreateElement("CreateInstanceRq")
		xmlRoot.AppendChild (nodeCreateInstanceRq)
		nodeCreateInstanceRq.SetAttribute "id", "e7"
		nodeCreateInstanceRq.SetAttribute "wfpNs", "ch.bw.wf.changemgmt.columbus_chgdevice"
		nodeCreateInstanceRq.SetAttribute "wfmNs", "Columbus Changemanagement"
		nodeCreateInstanceRq.SetAttribute "sessionId", "s1"
		
		'Das Node ObserverKey hinzufügen
		Dim nodeObserverKey : Set nodeObserverKey = pDict("XMLDocument").CreateElement("ObserverKey")
		nodeCreateInstanceRq.AppendChild (nodeObserverKey)
		nodeObserverKey.Text = pDict("ObserverKey")

		'Das Container Node ContextData hinzufügen
		Dim nodeContextData : Set nodeContextData = pDict("XMLDocument").CreateElement("ContextData")
		nodeCreateInstanceRq.AppendChild (nodeContextData)
		
		'Das Container Node AddDeviceActualParams hinzufügen
		Dim nodeChgDeviceActualParams : Set nodeChgDeviceActualParams = pDict("XMLDocument").CreateElement("dt:ChangeDeviceActualParams")
		nodeContextData.AppendChild (nodeChgDeviceActualParams)
		
		'Das Container Node DeviceIdentification hinzufügen
		Dim nodeDeviceIdentification : Set nodeDeviceIdentification = pDict("XMLDocument").CreateElement("dt:DeviceIdentification")
		nodeChgDeviceActualParams.AppendChild (nodeDeviceIdentification)
		
		'Das Node DeviceName hinzufügen
		Dim nodeDeviceName : Set nodeDeviceName = pDict("XMLDocument").CreateElement("dt:DeviceName")
		nodeDeviceIdentification.AppendChild (nodeDeviceName)
		nodeDeviceName.Text = pDict("DeviceName")
		
		'Das Node Domain hinzufügen
		Dim nodeDomain : Set nodeDomain = pDict("XMLDocument").CreateElement("dt:Domain")
		nodeDeviceIdentification.AppendChild (nodeDomain)
		nodeDomain.Text = pDict("Domain")

		'Das Node CompanyName hinzufügen
		Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
		nodeChgDeviceActualParams.AppendChild (nodeCmpyName)
		nodeCmpyName.Text = pDict("CompanyName")
		
		'Das Node CostCenter hinzufügen
		Dim nodeCostCenter : Set nodeCostCenter = pDict("XMLDocument").CreateElement("dt:CostCenter")
		nodeChgDeviceActualParams.AppendChild (nodeCostCenter)
		nodeCostCenter.Text = pDict("CostCenter")
		
		'Das Node MACAdess hinzufügen
		Dim nodeMACAddress : Set nodeMACAddress = pDict("XMLDocument").CreateElement("dt:MACAddress")
		nodeChgDeviceActualParams.AppendChild (nodeMACAddress)
		nodeMACAddress.Text = pDict("MACAddress")
		
		'Das Node SubnetMask hinzufügen
		Dim nodeSubnetMask : Set nodeSubnetMask = pDict("XMLDocument").CreateElement("dt:SubnetMask")
		nodeChgDeviceActualParams.AppendChild (nodeSubnetMask)
		nodeSubnetMask.Text = pDict("SubnetMask")
		
		'Das Node HwTypeId hinzufügen
		Dim nodeHWType : Set nodeHWType = pDict("XMLDocument").CreateElement("dt:HwTypeId")
		nodeChgDeviceActualParams.AppendChild (nodeHWType)
		nodeHWType.Text = pDict("HwTypeId")
		
		'Das Node OsTypeId hinzufügen
		Dim nodeOSType : Set nodeOSType = pDict("XMLDocument").CreateElement("dt:OsTypeId")
		nodeChgDeviceActualParams.AppendChild (nodeOSType)
		nodeOSType.Text = pDict("OsTypeId")
		
		'Das Node ActivationState hinzufügen
		Dim nodeActState : Set nodeActState = pDict("XMLDocument").CreateElement("dt:ActivationState")
		nodeChgDeviceActualParams.AppendChild (nodeActState)
		nodeActState.Text = pDict("ActivationState")

End Function
'---------------------------------------------------------------------
Function MIG_CreateDELXML2Columbus (hlSrvContext, pDict)
						
		'Root Element aus dem XML ermitteln.
		Dim xmlRoot : Set xmlRoot = pDict("XMLDocument").DocumentElement
		
		'Das Node CreateInstanceReq hinzufügen
		Dim nodeCreateInstanceRq : Set nodeCreateInstanceRq = pDict("XMLDocument").CreateElement("CreateInstanceRq")
		xmlRoot.AppendChild (nodeCreateInstanceRq)
		nodeCreateInstanceRq.SetAttribute "id", "e7"
		nodeCreateInstanceRq.SetAttribute "wfpNs", "ch.bw.wf.changemgmt.columbus_removedevice"
		nodeCreateInstanceRq.SetAttribute "wfmNs", "Columbus Changemanagement"
		nodeCreateInstanceRq.SetAttribute "sessionId", "s1"
		
		'Das Node ObserverKey hinzufügen
		Dim nodeObserverKey : Set nodeObserverKey = pDict("XMLDocument").CreateElement("ObserverKey")
		nodeCreateInstanceRq.AppendChild (nodeObserverKey)
		nodeObserverKey.Text = pDict("ObserverKey")

		'Das Container Node ContextData hinzufügen
		Dim nodeContextData : Set nodeContextData = pDict("XMLDocument").CreateElement("ContextData")
		nodeCreateInstanceRq.AppendChild (nodeContextData)
		
		'Das Container Node AddDeviceActualParams hinzufügen
		Dim nodeRemoveDeviceActualParams : Set nodeRemoveDeviceActualParams = pDict("XMLDocument").CreateElement("dt:RemoveDeviceActualParams")
		nodeContextData.AppendChild (nodeRemoveDeviceActualParams)
		
		'Das Container Node DeviceIdentification hinzufügen
		Dim nodeDeviceIdentification : Set nodeDeviceIdentification = pDict("XMLDocument").CreateElement("dt:DeviceIdentification")
		nodeRemoveDeviceActualParams.AppendChild (nodeDeviceIdentification)
		
		'Das Node DeviceName hinzufügen
		Dim nodeDeviceName : Set nodeDeviceName = pDict("XMLDocument").CreateElement("dt:DeviceName")
		nodeDeviceIdentification.AppendChild (nodeDeviceName)
		nodeDeviceName.Text = pDict("DeviceName")
		
		'Das Node CompanyName hinzufügen
		'Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
		'nodeDeviceIdentification.AppendChild (nodeCmpyName)
		'nodeCmpyName.Text = pDict("CompanyName")
		
		'Das Node Domain hinzufügen
		Dim nodeDomain : Set nodeDomain = pDict("XMLDocument").CreateElement("dt:Domain")
		nodeDeviceIdentification.AppendChild (nodeDomain)
		nodeDomain.Text = pDict("Domain")
				
End Function
'----------------------------------------------------------------------------------------------------------
'Wenn beide Werte ein Datum sind, muss geprüft werden ob das Enddatum nach dem
'Start Datum liegt. Falls nicht wird "False" zurückgegeben.
Function MigCheckDatePeriod(hlContext, StartDate, EndDate)
	MigCheckDatePeriod= False
      
	If DatePart("d", CDate(StartDate)) <> "0" Then
		If DatePart("d", CDate(StartDate)) < DatePart("d", CDate(EndDate)) Then
			MigCheckDatePeriod= False
		Else
			MigCheckDatePeriod= True
		End If

		If DatePart("yyyy", CDate(StartDate)) > DatePart("yyyy", CDate(EndDate)) Then
			MigCheckDatePeriod= False
		Else
			If DatePart("y", CDate(StartDate)) > DatePart("y", CDate(EndDate)) Then
				If DatePart("yyyy", CDate(StartDate)) < DatePart("yyyy", CDate(EndDate)) Then
					MigCheckDatePeriod= True
				Else
					MigCheckDatePeriod= False
				End If
			Else
				MigCheckDatePeriod= True
			End If
		End If
	End If
End Function
'---------------------------------------------------------------------
Function MIG_CheckCostCenter(hlSrvContext,strCostCenter)
		MIG_CheckCostCenter = False
		
		Dim srchQuery : srchQuery = ""
		srchQuery = "SEARCH Division WHERE OrganizationBilling.CostCenter_CA.CostCenter = """ & strCostCenter & """"
		Dim Qry : Set Qry = Nothing
		Dim rsltQuery : rsltQuery = ""
		Set Qry = hlSrvContext.OpenSearch (srchQuery)
		rsltQuery = Qry.GetItems(0, -1, -1, 0)
		If uBound(rsltQuery) >= 0 Then
			MIG_CheckCostCenter = True
		End If
		
End Function
'---------------------------------------------------------------------
' Check whether the agent (contact) is allowed to make 
' changes/modifications/create new entities of any objectdefinition based
' on the InternalMIGPartnerID of the contact and the MIGPartnerID of the object

Function CheckAgentHasMIGPartnerID(hlContext, relObjMIGPartnerID) 'BOOL

	Dim flagAuthorized : flagAuthorized = False
	Dim intAgentID : intAgentID = hlContext.GetAgentID()
	Dim objPerson : Set objPerson = Nothing
	Dim strPersonInternalMIGPartnerIDs

	Set objPerson = hlContext.GetPersonOfAgent(intAgentID)

	If IsHLObject(hlContext, objPerson) = True Then
		
		If relObjMIGPartnerID <> "" Then
	
			strPersonInternalMIGPartnerIDs = objPerson.GetValue("MIGAgentInformation.InternalMIGPartnerID",0,0,0,0)		

				If InStr(strPersonInternalMIGPartnerIDs,relObjMIGPartnerID) > 0 Then
					flagAuthorized = True
				End If
		Else
			'If relObjMIGPartnerID is Null or empty, modification allowed
			flagAuthorized = True
		End If

	End If

	'return
	CheckAgentHasMIGPartnerID = flagAuthorized
End Function
'---------------------------------------------------------------------