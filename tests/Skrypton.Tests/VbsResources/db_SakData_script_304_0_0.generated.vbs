'---------------------------------------------------------------
'Diese Funktion ermittelt den Standard-Eintrag zum angegebenen Attribut aus
'dem Dictionary.
'Wenn der Parameter "GetAll" auf False steht wird als Rueckgabewert fuer die Funktion
'ebenfalls "False" ausgegben, wenn mehr als ein Standardeintrag gefunden wird.
'Wenn fuer den Parameter "True" angeben wird, prueft die Funktion ob es tatsaechlich
'nur einen Standard-Eintrag gibt, sonst "False".
Public Function GetCommunicationDefault(ByRef hlContext, ByRef hlObject, ByRef dict, ByRef GetAll)
  GetCommunicationDefault = False
  Dim ItemCount
  ItemCount = 0
  Dim strValue
  strValue = ""

  Dim ItemIDs
  ItemIDs = ""
  ItemIDs = hlObject.GetContentIDs(dict("Compound"), 0)

  Dim Item
  Item = 0
  For Each Item In ItemIDs
    Dim defItem
    defItem = False
    defItem = GetFlagValue(hlContext, hlObject, dict("Default"), Item, 0)
    IF CBool(defItem) = True THEN
      ItemCount = ItemCount + 1
      strValue = hlObject.GetValue(dict("Value"), 0, Item, 0, 0)
      IF CBool(GetAll) = False THEN
        Exit For
      END IF
    END IF
  Next
  IF ItemCount > 1 THEN
    GetCommunicationDefault = False
    Exit Function
  ELSE
    GetCommunicationDefault = True
    dict("DefValue") = strValue
  END IF
End Function
'----------------------------------------------------------------------------------------------------------
'Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
Public Sub Trace(ByRef hlContext, ByRef text)
  hlContext.trace 1, text
End Sub
'---------------------------------------------------------------
'Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
Public Sub SetPersonInformation(ByRef hlContext, ByRef hlObject, ByRef dict)
  'Aus dem Dictionary wird das Attribut und der dazugehoerige Wert ermittelt.
  Dim AttrDef
  AttrDef = ""
  AttrDef = "PersonInformation." & dict("PersInfoAttr")

  Dim strAttrValue
  strAttrValue = ""
  strAttrValue = dict("DefValue")

  hlObject.SetValue AttrDef, 0, 0, 0, strAttrValue
End Sub
'---------------------------------------------------------------
Public Function IsHLObject(ByRef hlContext, ByRef hlObject)
  '	Trace hlContext, "IsObject " & IsObject(hlObject)
  '	Trace hlContext, "IsNull " & IsNull(hlObject)
  '	Trace hlContext, "IsEmpty " & IsEmpty(hlObject)
  '	Trace hlContext, "Leerstring "
  '	Trace hlContext, "Leerstring " & hlObject = ""
  Trace hlContext, "Type " & VarType(hlObject)
  IsHLObject =((IsObject(hlObject) = True) And((hlObject Is Nothing) = False))
End Function
'-------------------------------------------------------------------
Public Function GetBaseType(ByRef hlContext, ByRef hlObject)
  GetBaseType = hlObject.GetValue("HLOBJECTINFO.BASETYPE", 0, 0, 0, 0)
End Function
'---------------------------------------------------------------
'Dies ist eine rekursive Function zum ermitteln der Organisationshierarchie,
'ausgehend vom der ersten OU ueberhalb einer Person.
'Die Variable "strOrgUnits" ist der Out-Parameter der Function.
Public Function GetPersonOrganisation(ByRef hlContext, ByRef hlOrgUnit, ByRef strOrgUnits)
  GetPersonOrganisation = 0
  Dim retval
  retval = 0

  'Wenn noch keine OU ermittelt wurde, wird der Name der ersten OU eingetragen.
  'Andernfalls, wird jede weitere OU einfach angehangen.
  IF strOrgUnits = "" THEN
    strOrgUnits = hlOrgUnit.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
  ELSE
    strOrgUnits = strOrgUnits & ", " & hlOrgUnit.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
  END IF

  'Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
  'fuer die naechste Abfrage gewaehlt werden kann.
  Dim NextOrgUnit
  Dim orgaType
  orgaType = ""
  orgaType = hlOrgUnit.GetType
  IF orgaType = "Division" THEN
    NextOrgUnit = hlOrgUnit.GetItems(65536, 0, 0, "CompanyView")
  END IF
  IF orgaType = "Site" THEN
    NextOrgUnit = hlOrgUnit.GetItems(65536, 0, 0, "Site2Company")
  END IF
  IF orgaType = "Company" THEN
    NextOrgUnit = hlOrgUnit.GetItems(65536, 0, 0, "Company2Company")
  END IF

  'Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
  'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
  IF IsArray(NextOrgUnit) THEN
    IF UBound(NextOrgUnit) > = 0 THEN
      retval = GetPersonOrganisation(hlContext, NextOrgUnit(0), strOrgUnits)
    ELSE
      Exit Function
    END IF
  END IF
End Function
'---------------------------------------------------------------
'ueber diese Function wird fuer ein Flag Attribut immer der Wert
'True oder False ausgegeben.
Public Function GetFlagValue(ByRef hlContext, ByRef hlObject, ByRef hlattribute, ByRef hlcontentid, ByRef hlsuid)
  GetFlagValue = hlObject.GetValue(hlattribute, 0, hlcontentid, hlsuid, 0)
  IF GetFlagValue = "" THEN
    GetFlagValue = False
  END IF
End Function
'-------------------------------------------------------------------
'Diese Function ermitellt eine Fehlermeldung aus dem helpLine
'Woerterbuch ohne Parameter.
Public Function GetErrMsg0(ByRef hlContext, ByRef LocaleID, ByRef ErrCode)
  GetErrMsg0 = ""

  Dim strErrMsg
  strErrMsg = ""
  strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
  strErrMsg = strErrMsg & vbNewLine & "(Code: " & ErrCode & ")"

  'Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
  'Rueckgabewert der Function ist die Fehlermeldung.
  GetErrMsg0 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
'---------------------------------------------------------------

'Das Script ermittelt auf Basis der ersten uebergeordneten OU den gesamten Pfad bis zur Firma oder Konzern
'und speichert diesen in das Hilfsattribut PersonInformation.PersonOrganisation.
'This script detects the entire path based on the first parent OU up to the company or holding
'and saves them into the attribute PersonInformation.PersonOrganisation.
Public Sub SetPersonOrganization(ByRef hlContext, ByRef hlPerson, ByRef dict)
  Dim FirstOrgUnit
  Set FirstOrgUnit = Nothing
  Set FirstOrgUnit = hlContext.GetRelatedObject

  IF IsHLObject(hlContext, FirstOrgUnit) = True THEN
    IF FirstOrgUnit.GetType < > "Company" And FirstOrgUnit.GetType < > "Division" THEN
      Set FirstOrgUnit = Nothing
    END IF
  END IF

  IF IsHLObject(hlContext, FirstOrgUnit) = False THEN
    Dim rsltOrgUnit
    rsltOrgUnit = ""
    rsltOrgUnit = hlPerson.GetItems(65536, 0, 0, "Person2Organization")
    IF UBound(rsltOrgUnit) > = 0 THEN
      Set FirstOrgUnit = rsltOrgUnit(0)
    END IF
  END IF

  IF IsHLObject(hlContext, FirstOrgUnit) = True THEN
    IF GetBaseType(hlContext, FirstOrgUnit) = "ORGANISATION" THEN
      Dim retval
      retval = ""
      Dim strOrgUnits
      strOrgUnits = ""
      retval = GetPersonOrganisation(hlContext, FirstOrgUnit, strOrgUnits)

      dict("DefValue") = strOrgUnits
      dict("PersInfoAttr") = "PersonOrganization"
      Call SetPersonInformation(hlContext, hlPerson, dict)
    END IF
  END IF
End Sub
'---------------------------------------------------------------
'SACM
'----------------------------------------------------------------------------------------------------------
'Globale Konstanten fuer freie Assoziationsdefinitionen
Const Skrypton.LegacyParser.Tokens.Basic.NameToken:HLASC_SoftwareLicenseFolderView = Skrypton.LegacyParser.Tokens.Basic.StringToken:LicenseFolderView

Const Skrypton.LegacyParser.Tokens.Basic.NameToken:HLASC_SoftwareLicenseGroupView = Skrypton.LegacyParser.Tokens.Basic.StringToken:LicenseGroupView

Const Skrypton.LegacyParser.Tokens.Basic.NameToken:HLASC_Software2Computer = Skrypton.LegacyParser.Tokens.Basic.StringToken:Software2Computer

'----------------------------------------------------------------------------------------------------------
'Prozedur fuellt die Umzugshistorie fuer das entsprechende Objekt
Public Sub SetAssetHistory(ByRef hlContext, ByRef hlObjectA, ByRef hlObjectB, ByRef created)

  Dim productDefName
  productDefName = hlObjectB.GetType()

  IF (productDefName < > "Software" And productDefName < > "SoftwareLicence") THEN
    Dim agentID, contentID, personOfAgent, personName, orgUnitName
    contentID = hlObjectB.GenerateContentID()
    agentID = hlContext.GetAgentID()
    orgUnitName = hlObjectA.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
    Set personOfAgent = hlContext.GetPersonOfAgent(agentID)
    IF (personOfAgent Is Nothing) THEN
      Dim strErrMsg
      strErrMsg = GetErrMsg0(hlContext, hlContext.GetLocaleID, "#ERR_SETASSETHISTORY")
      Trace hlContext, strErrMsg
      'hlContext.abortcommand strErrMsg
    ELSE
      personName = personOfAgent.GetValue("PersonGeneral.Name", 0, 0, 0, 0)
      personName = personName & ", "
      personName = personName & personOfAgent.GetValue("PersonGeneral.GivenName", 0, 0, 0, 0)
    END IF
    hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangedBy", 0, contentID, 0, personName
    hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID", 0, contentID, 0, agentID
    hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryChangeDate", 0, contentID, 0, Now()
    hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryOrgUnit", 0, contentID, 0, orgUnitName
    hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryOrgUnitID", 0, contentID, 0, hlObjectA.GetID()

    IF (created = True) THEN
      hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryAction", 0, contentID, 0, "HistoryActionCreated"
    ELSE
      hlObjectB.SetValue "AssocHistory.HistoryInformation_CA.HistoryAction", 0, contentID, 0, "HistoryActionDeleted"
    END IF
  END IF
End Sub
'---------------------------------------------------------------
'Diese Function ermitellt eine Fehlermeldung aus dem helpLine
'Woerterbuch mit einem Parameter.
Public Function GetErrMsg1(ByRef hlContext, ByRef LocaleID, ByRef ErrCode, ByRef Arg1)
  GetErrMsg1 = ""

  Dim strErrMsg
  strErrMsg = ""
  strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
  strErrMsg = Replace(strErrMsg, "%1", Arg1)
  strErrMsg = strErrMsg & vbLf & "(Code: " & ErrCode & ")"

  'Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
  'Rueckgabewert der Function ist die Fehlermeldung.
  GetErrMsg1 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
Public Function GetErrMsg2(ByRef hlContext, ByRef LocaleID, ByRef ErrCode, ByRef Arg1, ByRef Arg2)
  GetErrMsg2 = ""

  Dim strErrMsg
  strErrMsg = ""
  strErrMsg = hlContext.GetTranslation(ErrCode, LocaleID)
  strErrMsg = Replace(strErrMsg, "%1", Arg1)
  strErrMsg = Replace(strErrMsg, "%2", Arg2)
  strErrMsg = strErrMsg & vbLf & "(Code: " & ErrCode & ")"

  'Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
  'Rueckgabewert der Function ist die Fehlermeldung.
  GetErrMsg2 = Replace(strErrMsg, "%LF%", vbNewLine)
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Funktion wird geprueft, ob es unterhalb einer Software Suite
'bereits Lizenzumschlaege mit Lizenzen gibt.
Public Function GetReferenceLicenseCount(ByRef hlContext, ByRef hlSWFolder, ByRef chkFolderOnly, ByRef HLASC_SoftwareLicenseFolderView)
  GetReferenceLicenseCount = 0

  Dim rsltSWFolders
  rsltSWFolders = ""
  Dim SoftwareLicense
  Set SoftwareLicense = Nothing
  Dim objType
  objType = ""

  'Pruefen ob es Software Lizenzobjekte/Lizenzumschlaege unterhalb des Folders gibt.
  rsltSWFolders = hlSWFolder.GetItems(0, - 1, - 1, HLASC_SoftwareLicenseFolderView)

  For Each SoftwareLicense In rsltSWFolders
    objType = SoftwareLicense.GetType()
    IF objType = "LicenseFolder" THEN
      GetReferenceLicenseCount = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
      IF GetReferenceLicenseCount > 0 THEN
        Exit Function
      END IF
    END IF
    IF objType = "SoftwareLicense" And CBool(chkFolderOnly) = False THEN
      GetReferenceLicenseCount = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
      IF GetReferenceLicenseCount > 0 THEN
        Exit Function
      END IF
    END IF
  Next
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
'den obersten Lizenz Umschlag ermittelt. Auf dem Weg dort hin wird geprueft ob einer
'der Lizenzumschlaege eine Software Suite ist.
Public Function CheckForSoftwareSuiteFolder(ByRef hlContext, ByRef hlParentSWFolder, ByRef pDict, ByRef HLASC_SoftwareLicenseFolderView)
  CheckForSoftwareSuiteFolder = ""
  Dim retval
  retval = 0
  Dim NextSWFolder
  NextSWFolder = ""

  'Festhalten auf welcher Ebene ggf. eine Software Suite oberhalb des
  'Start Folders existiert. Die Variable muss von aussen mit einem Startwert
  'initialisiert werden.
  IF pDict("SoftwareSuiteFolderLevel") = 0 Or pDict("SoftwareSuiteFolderLevel") = "" THEN
    pDict("SoftwareSuiteFolderLevel") = 1
  ELSE
    pDict("SoftwareSuiteFolderLevel") = pDict("SoftwareSuiteFolderLevel") + 1
  END IF

  'Amhand des Flags "Software Suite" festellen ob ein Lizenzumschlag als Software Suite
  'gekennzeichnet ist. Falls Ja, Name des Umschlags auslesen und Funktion abbrechen.
  Dim CheckSoftwareSuite
  CheckSoftwareSuite = False
  CheckSoftwareSuite = GetFlagValue(hlContext, hlParentSWFolder, "SoftwareLicenseFolderDetail.FlagSoftwareSuite", 0, 0)
  IF CBool(CheckSoftwareSuite) = True THEN
    pDict("SoftwareSuiteFolder") = hlParentSWFolder.GetValue("OrganizationGeneral.Name", 0, 0, 0, 0)
    Exit Function
  END IF

  'Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
  'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
  NextSWFolder = hlParentSWFolder.GetItems(65536, - 1, - 1, HLASC_SoftwareLicenseFolderView)
  IF UBound(NextSWFolder) > = 0 THEN
    retval = CheckForSoftwareSuiteFolder(hlContext, NextSWFolder(0), pDict, HLASC_SoftwareLicenseFolderView)
  ELSE
    Exit Function
  END IF
End Function
'----------------------------------------------------------------------------------------------------------
'In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
'den obersten Lizenz Umschlag ermittelt und neu berechnet hat.
Public Function SetLicenseCounter(ByRef hlContext, ByRef hlSWFolder, ByRef pDict, ByRef assocName)
  SetLicenseCounter = 0
  Dim retval
  retval = 0

  'Dictionary Eintraege initalisieren
  pDict("SoftwareLicenses") = ""
  pDict("SumRefLicCounter") = 0
  pDict("SumInstLicCounter") = 0
  pDict("SumFreeLicCounter") = 0

  'Pruefen ob es Software Lizenzobjekte unterhalb des Folders gibt.
  pDict("SoftwareLicenses") = hlSWFolder.GetItems(0, - 1, - 1, assocName)

  'Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
  'Objekte gezaehlt werden muessen
  Dim CheckSoftwareSuite
  CheckSoftwareSuite = False
  CheckSoftwareSuite = GetFlagValue(hlContext, hlSWFolder, "SoftwareLicenseFolderDetail.FlagSoftwareSuite", 0, 0)

  IF UBound(pDict("SoftwareLicenses")) > = 0 THEN
    IF CBool(CheckSoftwareSuite) = False THEN
      Call CalcAllLicCounter(hlContext, pDict)
    ELSE
      Call CalcFolderLicCounter(hlContext, pDict)
    END IF
  END IF
  'Gesatmzahl der Lizenzen in den Lizenzumschlag zurueckschreiben
  hlSWFolder.SetValue "SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, pDict("SumRefLicCounter")
  hlSWFolder.SetValue "SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, pDict("SumInstLicCounter")

  'Wenn die Lizenzkontrolle durch den Applikations Server erfolgt ("Lizenzkontrolle durch Server")
  'dann die Anzahl freier Lizenzen immer auf den Wert "0" setzen.
  Dim CheckLicContrByServer
  CheckLicContrByServer = False
  CheckLicContrByServer = GetFlagValue(hlContext, hlSWFolder, "SoftwareLicenseFolderDetail.FlagLicenseControlledByServer", 0, 0)
  IF CBool(CheckLicContrByServer) = True THEN
    pDict("SumFreeLicCounter") = 0
  END IF
  hlSWFolder.SetValue "SoftwareLicenseCounter.FreeLicenseCount", 0, 0, 0, pDict("SumFreeLicCounter")

  'Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
  'fuer die naechste Abfrage gewaehlt werden kann.
  Dim NextSWFolder
  NextSWFolder = ""
  Dim a
  a = ""
  a = hlSWFolder.GetType
  IF a = "LicenseFolder" THEN
    NextSWFolder = hlSWFolder.GetItems(65536, 0, 0, assocName)
  END IF
  'Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
  'dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
  IF UBound(NextSWFolder) > = 0 THEN
    retval = SetLicenseCounter(hlContext, NextSWFolder(0), pDict, assocName)
  ELSE
    Exit Function
  END IF
End Function
'----------------------------------------------------------------------------------------------------------
Public Function IsValidObject(ByRef obj)
  IsValidObject =(IsObject(obj) And(Not(obj Is Nothing)))
End Function
'----------------------------------------------------------------------------------------------------------
Public Sub CalcAllLicCounter(ByRef hlContext, ByRef pDict)
  Dim SWRefLicCounter
  SWRefLicCounter = 0
  Dim SWInstCounter
  SWInstCounter = 0
  Dim SoftwareLicense
  Set SoftwareLicense = Nothing
  Dim objType
  objType = ""
  Dim lstLicStatus
  lstLicStatus = ""

  For Each SoftwareLicense In pDict("SoftwareLicenses")
    objType = SoftwareLicense.GetType()
    IF objType = "SoftwareLicense" THEN
      lstLicStatus = SoftwareLicense.GetValue("SoftwareLicenseDetail.LicenseStatus", 0, 0, 0, 0)
      IF lstLicStatus = "LicenseStatusValid" THEN
        SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
        pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
      END IF
    ELSE
      IF objType = "LicenseFolder" Or objType = "Software" THEN
        SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
        pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
        SWInstCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, 0))
        pDict("SumInstLicCounter") = pDict("SumInstLicCounter") + SWInstCounter
      END IF
    END IF
  Next
  'Anzahl freier Lizenzen berechnen und in den Folder schreiben.
  pDict("SumFreeLicCounter") = pDict("SumRefLicCounter") - pDict("SumInstLicCounter")

End Sub
'----------------------------------------------------------------------------------------------------------
Public Sub CalcFolderLicCounter(ByRef hlContext, ByRef pDict)

  Dim SWRefLicCounter
  SWRefLicCounter = 0
  Dim SWInstCounter
  SWInstCounter = 0
  Dim SoftwareLicense
  Set SoftwareLicense = Nothing
  Dim objType
  objType = ""
  Dim lstLicStatus
  lstLicStatus = ""

  For Each SoftwareLicense In pDict("SoftwareLicenses")
    objType = SoftwareLicense.GetType()
    IF objType = "LicenseFolder" Or objType = "Software" THEN
      SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
      pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter

      SWInstCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.InstalledLicenseCount", 0, 0, 0, 0))
      IF SWInstCounter > pDict("SumInstLicCounter") THEN
        pDict("SumInstLicCounter") = SWInstCounter
      END IF
    END IF
    IF objType = "SoftwareLicense" THEN
      lstLicStatus = SoftwareLicense.GetValue("SoftwareLicenseDetail.LicenseStatus", 0, 0, 0, 0)
      IF lstLicStatus = "LicenseStatusValid" THEN
        SWRefLicCounter = CheckIntegerValue(hlContext, SoftwareLicense.GetValue("SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, 0))
        pDict("SumRefLicCounter") = pDict("SumRefLicCounter") + SWRefLicCounter
      END IF
    END IF
  Next
  'Anzahl freier Lizenzen berechnen und in den Folder schreiben.
  pDict("SumFreeLicCounter") = pDict("SumRefLicCounter") - pDict("SumInstLicCounter")
End Sub
'----------------------------------------------------------------------------------------------------------
'Diese Function ueberprueft den ganzzahligen Wert (Integer).
Public Function CheckIntegerValue(ByRef hlContext, ByRef intval)
  IF intval = "" Or IsNumeric(intval) = False THEN
    CheckIntegerValue = 0
  ELSE
    CheckIntegerValue = CLng(intval)
  END IF
End Function
'----------------------------------------------------------------------------------------------------------
Public Function OnCreate_HasAssociationToDelete(ByRef hlContext, ByRef AscDefName, ByRef hlObjB)
  Dim result
  result = False
  Dim cAssociationChanges
  cAssociationChanges = 0
  cAssociationChanges = hlContext.GetAssociationChangesCount

  Dim oAssociationChange
  Set oAssociationChange = Nothing
  Dim AscDefNameChange
  AscDefNameChange = ""
  Dim ixAC
  ixAC = 0

  For ixAC = 0 To cAssociationChanges - 1
    Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)

    AscDefNameChange = oAssociationChange.AssociationType

    IF oAssociationChange.IsToDelete THEN
      IF (AscDefNameChange = AscDefName) THEN
        IF (hlObjB.GetID = oAssociationChange.EndB.GetID) THEN
          result = True
          Exit For
        END IF
        'check the ids
      END IF
      ' check the defnames
    END IF
    ' is to create
  Next
  OnCreate_HasAssociationToDelete = result
End Function
'----------------------------------------------------------------------------------------------------------
Public Function OnCreate_HasAssociationToCreate(ByRef hlContext, ByRef AscDefName, ByRef hlObjB)
  Dim result
  result = False
  Dim cAssociationChanges
  cAssociationChanges = 0
  cAssociationChanges = hlContext.GetAssociationChangesCount

  Dim oAssociationChange
  Set oAssociationChange = Nothing
  Dim AscDefNameChange
  AscDefNameChange = ""
  Dim ixAC
  ixAC = 0

  For ixAC = 0 To cAssociationChanges - 1
    Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)

    AscDefNameChange = oAssociationChange.AssociationType

    IF oAssociationChange.IsToCreate THEN
      IF (AscDefNameChange = AscDefName) THEN
        IF (hlObjB.GetID = oAssociationChange.EndB.GetID) THEN
          result = True
          Exit For
        END IF
        'check the ids
      END IF
      ' check the defnames
    END IF
    ' is to create
  Next
  OnCreate_HasAssociationToCreate = result
End Function

Public Function OnDelete_HasAssociationToCreate(ByRef hlContext, ByRef AscDefName, ByRef hlObjB)
  ' bool
  Dim result
  result = False

  'Anzahl der zu erstellenden oder loeschenden Assoziationen
  Dim cAssociationChanges
  cAssociationChanges = 0
  cAssociationChanges = hlContext.GetAssociationChangesCount

  Dim oAssociationChange
  Set oAssociationChange = Nothing
  Dim AscDefNameChange
  AscDefNameChange = ""
  Dim ixAC
  ixAC = 0

  For ixAC = 0 To cAssociationChanges - 1

    'Fuer jede Assoziations aenderung wird das entsprechende Infos (Objekt    ) ausgelsen.
    Set oAssociationChange = hlContext.GetAssociationChangeAt(ixAC)
    'Def Name der Assoc ermitteln, die angelegt werden soll
    AscDefNameChange = oAssociationChange.AssociationType

    IF oAssociationChange.IsToCreate THEN
      'ueberpruefen ob die gewuenschte Assoc auch angelegt werden soll.
      IF (AscDefNameChange = AscDefName) THEN
        IF (hlObjB.GetID = oAssociationChange.EndB.GetID) THEN
          result = True
          Exit For
        END IF
        'check the ids
      END IF
      ' check the defnames
    END IF
    ' is to create
  Next
  OnDelete_HasAssociationToCreate = result
End Function
'----------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------
Public Function GetAssociatedOrganizationalUnit(ByRef hlContext, ByRef lcid, ByRef hlChild, ByRef pDict, ByRef outParentDefName)
  GetAssociatedOrganizationalUnit = ""
  outParentDefName = ""

  Dim rsltParent
  rsltParent = ""
  rsltParent = hlChild.GetItems(65536, - 1, - 1, pDict("AssocID"))
  IF UBound(rsltParent) > = pDict("ParentCounter") THEN
    Dim objParent
    Set objParent = Nothing
    For Each objParent In rsltParent
      GetAssociatedOrganizationalUnit = objParent.GetValue(0, 0, 0, pDict("AttrName"), 0)
      outParentDefName = hlContext.GetDisplayName(objParent.GetValue(0, 0, 0, 0, "HLOBJECTINFO.DEFID"), lcid)
      Exit For
    Next
  END IF
End Function
'----------------------------------------------------------------------------------------------------------
