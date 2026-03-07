Public Sub IncReqOnLoad()
  Dim ReadOnly, NoPerson, NoAsset
  ReadOnly = True
  NoPerson = True
  NoAsset = True

  'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
  'First of all check whether the Case is write protected
  IF hlObj.IsReadOnly("CaseGeneral.Subject", 0) = 0 THEN
    ReadOnly = False
  END IF

  'Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
  'Check wether the Caller object exist
  IF IsObject(hlCaller) = True And EditSurname.Text <> "" THEN
    NoPerson = False
  END IF

  'VIP-Status des Anfragers abfragen und im Vorgang setzen
  Valid = hlCaller.HasContent("PersonGeneral.VIPLevel", 0, 0)
  IF Valid = 1 THEN
    VIP = hlCaller.GetValue("PersonGeneral.VIPLevel", 0, 0, 0, 0)
    'If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
    SELECT CASE vip
      CASE "VIPLevelVIP"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 1
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(142, 139, 254)
      CASE "VIPLevelITAdminDitzingen"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 2
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelITAdminTG"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 3
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelSAPKeyUserTUS"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 4
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelNon"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 0
        ComboVIPStatus.Disabled = True
        Person.BackColor = ""
    END SELECT
  END IF

  'Prüft ob ein Produkt Objekt vorhanden ist und ob dieses auch angezeigt wird
  'Check wether the Product object exist
  IF IsObject(hlProduct) = True And EditAssetModel.Text <> "" THEN
    NoAsset = False
  END IF

  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)

  'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check requester search status to set the caption of the button
  IF NoPerson = False THEN
    IF SearchCaller.GetSearchState = 3 THEN
      SearchCaller.Caption = "Reset"
    ELSE
      SearchCaller.Caption = "Betroffener"
    END IF
  END IF

  'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check Asset search status to set the caption of the button
  IF NoAsset = False THEN
    IF SearchAsset.GetSearchState = 3 THEN
      SearchAsset.Caption = "Reset"
    ELSE
      SearchAsset.Caption = "Inventar"
    END IF
  END IF

  IF NoAsset = False THEN
    'Setzen des Inventars
    'Setting the asset
    varString = ""
    varAType = hlProduct.GetType()
    IF varAType = "DesktopComputer" Or varAType = "ServerComputer" Or varAType = "NotebookComputer" Or varAType = "Printer" THEN
      IF EditHostname.Text <> "" THEN
        varString = EditHostname.Text
      END IF
      IF EditAssetModel.Text <> "" THEN
        varString = varString & " " & EditAssetModel.Text
      END IF
    ELSE
      IF EditAssetModel.Text <> "" THEN
        varString = EditAssetModel.Text
      ELSE
        EditAssetModel.Text = " "
      END IF
    END IF
    EditAssetModel.Text = varString
  END IF

  'Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
  Dim Anfrageart
  Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0)

  IF Anfrageart <> "RequestTypeIncident" THEN
    ComboImpact.Disabled = True
    ComboFunctionalRange.Disabled = True
  ELSE
    ComboImpact.Disabled = False
    ComboFunctionalRange.Disabled = False
  END IF

  IF Anfrageart <> "RequestTypeContact" THEN
    CaseProblem.Disabled = False
    ComboBoxEmailCaller.Disabled = False
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    ComboIncidentStatus.Disabled = False

  ELSE
    CaseProblem.Disabled = True
    ComboBoxEmailCaller.Disabled = True
    CaseDiagnosis.Disabled = True
    KeywordTree.Disabled = True
    Attachment.Disabled = True
    ComboProductionalRelevanz.Disabled = true
    ComboIncidentStatus.Disabled = True
  END IF

  'Zugriff auf Übersichts-Buttons regeln
  IF ReadOnly = False THEN
    ButtonShowOverView.Disabled = False
    ButtonEmailPreview.Disabled = False
    EditSubjectCase.Disabled = False
  ELSE
    ButtonShowOverView.Disabled = True
    ButtonEmailPreview.Disabled = True
    EditSubjectCase.Disabled = True
  END IF

  'Einfärben der GrupBox CaseAttributes je nach Priorität
  SELECT CASE hlObj.GetValue("CaseClassificationAttribute.Priority", 0, 0, 0, 0)
    CASE "Priority1"
      CaseAttributes.BackColor = RGB(107, 105, 248)
    CASE "Priority2"
      CaseAttributes.BackColor = RGB(119, 170, 251)
    CASE "Priority3"
      CaseAttributes.BackColor = RGB(132, 235, 255)
    CASE "Priority4"
      CaseAttributes.BackColor = RGB(128, 213, 177)
    CASE "Priority5"
      CaseAttributes.BackColor = RGB(123, 190, 99)
    CASE ELSE
      CaseAttributes.BackColor = RGB(248, 245, 240)
  END SELECT

  'Bei Status ToProof wird die Email-Tab angewählt
  IF hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0) = "IncidentStatusToProof" THEN
    TabPageEmail.UiActive = True
  ELSE
  END IF

End Sub
Public Sub OnSUIDAdded()
  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)

  Dim ReadOnly, NoPerson, NoAsset
  ReadOnly = True
  NoPerson = True
  NoAsset = True

  'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
  'First of all check whether the Case is write protected
  IF hlObj.IsReadOnly("CaseGeneral.Subject", 0) = 0 THEN
    ReadOnly = False
  END IF


  'Status auf "In Bearbeitung" setzen
  hlObj.SetValue "IncidentAttribute.IncidentStatus", 0, 0, 0, "IncidentStatusInProgress"

  'Wenn Vorgang erweitert wird, wird die Zuständigkeit des Agenten ermittelt und gestezt.
  Dim GetLastSUIdx
  GetLastSUIdx = 0
  Dim suindices
  suindices = hlobj.GetSvcUnitIndices()
  GetLastSUIdx = UBound(suindices)
  IF GetLastSUIdx > 0 THEN
    Dim agent
    agent = hlObj.GetValue("SUINFO.EDITOR", 0, 0, GetLastSUIdx + 1, 1)
    Dim person, helper, responsibilty
    Set helper = CreateObject("helpline.hlcontrols.HLHelperPFA")
    Set person = helper.GetPersonForAgent(model.GetClientContext, clng(agent))
    IF isObject(person) = True THEN
      responsibility = person.GetValue("PersonGeneralTrumpf.Responsibility", 0, 0, 0, 0)
      IF responsibility = "ResponsibilityBSZDitzingen" THEN
        hlObj.SetValue "IncidentAttribute.Responsibility", 0, 0, 0, "ResponsibilityBSZDitzingen"
      ELSE
        hlObj.SetValue "IncidentAttribute.Responsibility", 0, 0, 0, "ResponsibilityLocalIT"
      END IF
    END IF
  END IF


  'Zugriff auf Übersichts-Buttons regeln
  IF ReadOnly = False THEN
    ButtonShowOverView.Disabled = False
    ButtonEmailPreview.Disabled = False
    EditSubjectCase.Disabled = False
  ELSE
    ButtonShowOverView.Disabled = True
    ButtonEmailPreview.Disabled = True
    EditSubjectCase.Disabled = True
  END IF
  'Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
  Dim Anfrageart
  Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0)
  IF Anfrageart <> "RequestTypeContact" THEN
    ComboIncidentStatus.Disabled = False
  ELSE
    ComboIncidentStatus.Disabled = True
  END IF

  'Bei 2nd Level Dialog setzen der Benachrichtigung auf Email
  hlObj.SetValue "CaseGeneral.DefaultNotification", 0, 0, 0, "DefaultNotificationEmail"

End Sub
Public Sub SearchAsset_AfterExecute()
  'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check Asset search status to set the caption of the button
  IF SearchAsset.GetSearchState = 3 THEN
    SearchAsset.Caption = "Reset"
  ELSE
    SearchAsset.Caption = "Inventar"
  END IF

End Sub
Public Sub SearchAsset_AfterReset()
  Set objO = SearchAsset.GetObject("product", False)
  Set objT = SearchAsset.GetObject("product", True)

  Call objT.SetValue("AssetGeneral.AssetName", 0, 0, 0, "")
  Call objT.SetValue("AssetGeneral.Hostname", 0, 0, 0, "")
  Call objT.SetValue("TrumpfAssetGeneral.CINumber", 0, 0, 0, "")

  'Prüft ob Anfrager Objekt nicht vorhanden ist
  'Check wether the Caller object exist
  IF IsObject(hlCaller) = False Or hlCaller.objID = 0 THEN
    Call hlObj.SetValue("CaseGeneral.CostCenter", 0, 0, 0, "")
  END IF

  'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check Asset search status to set the caption of the button
  IF SearchAsset.GetSearchState = 3 THEN
    SearchAsset.Caption = "Reset"
  ELSE
    SearchAsset.Caption = "Inventar"
  END IF

End Sub
Public Sub SearchAsset_Click()
  Dim ReadOnly, NoProduct
  ReadOnly = True
  NoProduct = True

  'Wenn kein Inventar gefunden wurde, abbrechen
  'Cancel If no Asset was found
  IF hlProduct.GetType() = "TEMPOBJECT" THEN
    Exit Sub
  END IF

  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)

  'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
  'First of all check whether the Case is write protected
  IF hlObj.IsReadOnly("CaseGeneral.Subject", 0) = 0 THEN
    ReadOnly = False
  END IF

  'Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
  'Check wether the Caller object exist
  IF IsObject(hlProduct) = True And EditHostname.Text <> "" THEN
    NoProduct = False
  END IF

  IF ReadOnly = False THEN
    'Setzen des Inventars
    'Setting the asset
    varString = ""
    varAType = hlProduct.GetType()
    IF varAType = "DesktopComputer" Or varAType = "ServerComputer" Or varAType = "NotebookComputer" Or varAType = "Printer" THEN
      IF EditHostname.Text <> "" THEN
        varString = EditHostname.Text
      END IF
      IF EditAssetModel.Text <> "" THEN
        varString = varString & " " & EditAssetModel.Text
      END IF
    ELSE
      IF EditAssetModel.Text <> "" THEN
        varString = EditAssetModel.Text
      ELSE
        EditAssetModel.Text = " "
      END IF
    END IF
    EditAssetModel.Text = varString
  END IF

End Sub
Public Sub SearchCaller_AfterExecute()
  'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check requester search status to set the caption of the button
  IF SearchCaller.GetSearchState = 3 THEN
    SearchCaller.Caption = "Reset"
  ELSE
    SearchCaller.Caption = "Search"
  END IF

  'VIP-Status des Anfragers abfragen und Imp Vorgang setzen
  Valid = hlCaller.HasContent("PersonGeneral.VIPLevel", 0, 0)
  IF Valid = 1 THEN
    VIP = hlCaller.GetValue("PersonGeneral.VIPLevel", 0, 0, 0, 0)
    'If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
    SELECT CASE vip
      CASE "VIPLevelVIP"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 1
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(142, 139, 254)
      CASE "VIPLevelITAdminDitzingen"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 2
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelITAdminTG"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 3
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelSAPKeyUserTUS"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 4
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205, 250, 255)
      CASE "VIPLevelNon"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0, 0
        ComboVIPStatus.Disabled = True
        Person.BackColor = ""
    END SELECT
  END IF

  sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0)
  strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
  Dim tempmail
  tempmail = EditEmailAddress.text
  'Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
  'Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
  Dim strIncStatus
  strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
  strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
  strEmail = ""
  CallerCount = 0
  CallerCount = hlObj.GetItemCount(0, 130)

  IF CallerCount > 0 THEN
    Dim CaseCallers
    Set CaseCallers = Nothing
    CaseCallers = hlObj.GetItems(0, - 1, - 1, 130)
    For Each Caller In CaseCallers
      CallerType = Caller.GetType
      IF CallerType = "Employee" THEN
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
        IF mailadr <> "" THEN
          strEmail = strEmail + mailadr + ";"
        END IF
      END IF
    Next

  ELSE
    strEmail = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF

  IF InStr(strEmail, tempmail) > 0 THEN
  ELSE
    strEmail = tempmail + ";" + strEmail
  END IF

  IF strEmail = "" THEN
    strEmail = hlObj.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF
  IF strEmail = "-" THEN
    strEmail = ""
  END IF
  IF sendmail = "EmailCallerYes" THEN
    hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
    hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
    TextBoxEmailTo.Required = True
    TextBoxEmailSubject.Required = True
    GroupBoxEmail.Disabled = False
  ELSE
    hlObj.SetValue "EmailSUAttribute.EmailSearchName", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailSearchResult", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailCC", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT", 0, 0, 0, ""
    GroupBoxEmail.Disabled = True
    TextBoxEmailTo.Required = False
    TextBoxEmailSubject.Required = False
  END IF

End Sub
Public Sub SearchCaller_AfterReset()
  Set objO = SearchCaller.GetObject("caller", False)
  Set objT = SearchCaller.GetObject("caller", True)

  Call objT.SetValue("PersonGeneral.PersonSurname", 0, 0, 0, "")
  Call objT.SetValue("PersonGeneral.PersonGivenName", 0, 0, 0, "")
  Call objT.SetValue("PersonInformation.PersonOrganisation", 0, 0, 0, "")
  Call objT.SetValue("PersonInformation.PhoneNumber", 0, 0, 0, "")
  Call hlObj.SetValue("CaseGeneral.CostCenter", 0, 0, 0, "")


  'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
  'Check requester search status to set the caption of the button
  IF SearchCaller.GetSearchState = 3 THEN
    SearchCaller.Caption = "Reset"
  ELSE
    EditSurname.Text = ""
    SearchCaller.Caption = "Search"
  END IF

  'VIP-Status zurücksetzen
  ComboVIPStatus.SelectItem 0, 0
  Person.BackColor = RGB(248, 245, 240)

End Sub
Public Sub SearchCaller_Click()
  Dim ReadOnly
  ReadOnly = True

  'Wenn keine Person gefunden wurde, abbrechen
  'Cancel If no person was found
  IF hlCaller.GetType() = "TEMPOBJECT" THEN
    Exit Sub
  END IF

  'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
  'First of all check whether the Case is write protected
  IF hlObj.IsReadOnly("CaseGeneral.Subject", 0) = 0 THEN
    ReadOnly = False
  END IF

  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)





End Sub
Public Sub SetProblemText2Subject()
  varSubject = Left(EditProblem.Text, 100)
  IF EditSubjectCase.Text = "" THEN
    EditSubjectCase.Text = replace(varSubject, Chr(13) & Chr(10), " ")
  END IF

End Sub
Public Sub ComboIncidentStatus_SelectionChanged()
  Dim tempmail
  tempmail = EditEmailAddress.text
  Dim strIncStatus
  strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
  strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
  strEmail = ""
  CallerCount = 0
  CallerCount = hlObj.GetItemCount(0, 130)

  IF CallerCount > 0 THEN
    Dim CaseCallers
    Set CaseCallers = Nothing
    CaseCallers = hlObj.GetItems(0, - 1, - 1, 130)
    For Each Caller In CaseCallers
      CallerType = Caller.GetType
      IF CallerType = "Employee" THEN
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
        IF mailadr <> "" THEN
          strEmail = strEmail + mailadr + ";"
        END IF
      END IF
    Next
  ELSE
    strEmail = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF

  IF InStr(strEmail, tempmail) > 0 THEN
  ELSE
    strEmail = tempmail + ";" + strEmail
  END IF

  IF strEmail = "" THEN
    strEmail = hlObj.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF
  IF strEmail = "-" THEN
    strEmail = ""
  END IF
  SELECT CASE strIncStatus
    CASE "IncidentStatusSolved"
      hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerYes"
      hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
      hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
      hlObj.SetValue "SUINFO.PUBLISHED", 0, 0, 0, "1"
      GroupBoxEmail.Disabled = False
      LabelEmailBody.TextColor = "Red"
      ComplexTextEmailBody.Required = True
      TextBoxEmailTo.Required = True
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = True
    CASE "IncidentStatusClosed"
      hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
      hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
      hlObj.SetValue "SUINFO.PUBLISHED", 0, 0, 0, "1"
      GroupBoxEmail.Disabled = False
      LabelEmailBody.TextColor = "Red"
      ComplexTextEmailBody.Required = True
      IF strEmail = "" THEN
        TextBoxEmailTo.Required = False
        hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerNo"
      ELSE
        hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerYes"
        TextBoxEmailTo.Required = True
      END IF
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = True
    CASE "IncidentStatusTimephased"
      hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerYes"
      hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
      hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
      hlObj.SetValue "SUINFO.PUBLISHED", 0, 0, 0, "1"
      GroupBoxEmail.Disabled = False
      LabelEmailBody.TextColor = "Red"
      ComplexTextEmailBody.Required = True
      TextBoxEmailTo.Required = True
      EditResubmissionTime.Required = True
      EditResubmissionTime.Disabled = false
      ComboBoxEmailCaller.Disabled = True
    CASE "IncidentStatusWaitingforCustomer"
      hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerYes"
      hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
      hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
      hlObj.SetValue "SUINFO.PUBLISHED", 0, 0, 0, "1"
      GroupBoxEmail.Disabled = False
      LabelEmailBody.TextColor = "Red"
      ComplexTextEmailBody.Required = True
      TextBoxEmailTo.Required = True
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = True
    CASE "IncidentStatusWaitingforExtern"
      LabelEmailBody.TextColor = "Black"
      ComplexTextEmailBody.Required = False
      TextBoxEmailTo.Required = False
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = False
    CASE "IncidentStatusToProof"
      LabelEmailBody.TextColor = "Black"
      ComplexTextEmailBody.Required = False
      TextBoxEmailTo.Required = False
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = False
    CASE "IncidentStatusRouted"
      hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerNo"
      LabelEmailBody.TextColor = "Black"
      ComplexTextEmailBody.Required = False
      TextBoxEmailTo.Required = False
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = False
    CASE "IncidentStatusNew"
      LabelEmailBody.TextColor = "Black"
      ComplexTextEmailBody.Required = False
      TextBoxEmailTo.Required = False
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = False
    CASE "IncidentStatusInProgress"
      hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, "EmailCallerNo"
      LabelEmailBody.TextColor = "Black"
      ComplexTextEmailBody.Required = False
      TextBoxEmailTo.Required = False
      EditResubmissionTime.Required = False
      EditResubmissionTime.Disabled = true
      EditResubmissionTime.DeleteContent
      ComboBoxEmailCaller.Disabled = False
  END SELECT


End Sub
Public Sub ComboRequestType_SelectionChanged()
  Dim Anfrageart, Status
  Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0)
  Status = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)

  ComboProductionalRelevanz.Disabled = false

  IF Anfrageart <> "RequestTypeIncident" THEN
    ComboImpact.Disabled = True
    ComboFunctionalRange.Disabled = True
    hlObj.SetValue "CaseClassificationAttribute.Impact", 0, 0, 0, "ImpactOne"
    hlObj.SetValue "IncidentAttribute.FunctionalRange", 0, 0, 0, "FunctionalRangePartFailure"
    hlObj.SetValue "IncidentAttribute.ProductionalRelevanz", 0, 0, 0, "ProductionalRelevanzAdministrativeProcess"
  ELSE
    ComboImpact.Disabled = False
    ComboFunctionalRange.Disabled = False
    hlObj.SetValue "IncidentAttribute.ProductionalRelevanz", 0, 0, 0, "ProductionalRelevanzSupportProcess"
  END IF

  IF Anfrageart <> "RequestTypeContact" THEN
    CaseProblem.Disabled = False
    IF Status <> "IncidentStatusClosed" THEN
      ComboBoxEmailCaller.Disabled = False
    ELSE
      ComboBoxEmailCaller.Disabled = True
    END IF
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    CaseAttributes.Disabled = False
    ComboIncidentStatus.Disabled = False
  ELSE
    EditProblem.Text = ""
    CaseProblem.Disabled = True
    ComboBoxEmailCaller.Disabled = True
    EditDiagnosis.Text = ""
    CaseDiagnosis.Disabled = True
    KeywordTree.Disabled = True
    Attachment.Disabled = True
    CaseAttributes.Disabled = True
    ComboRequestType.Disabled = False
    ComboProductionalRelevanz.Disabled = true
    ComboIncidentStatus.Disabled = True
  END IF



End Sub
Public Sub OnSave()
  'Priorität leeren, damit globale SLA´s auch runterstufen können
  hlObj.SetValue "CaseClassificationAttribute.Priority", 0, 0, 0, "Priority5"

  CheckOverView = ""
  CheckOverView = hlObj.GetValue("CaseGeneral.Overview", 0, 0, 0, 0)
  IF CheckOverView <> "" THEN
    hlObj.SetValue "CaseGeneral.Overview", 0, 0, 0, ""
  END IF
  CheckSummaryHTML = ""
  CheckSummaryHTML = hlObj.GetValue("CaseGeneral.SummaryHTML.TEXTVALUE", 0, 0, 0, 0)
  IF CheckSummaryHTML <> "" THEN
    hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE", 0, 0, 0, ""
    hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT", 0, 0, 0, ""
    'Button "Übersicht" entsperren
    ButtonShowOverView.Disabled = False
  END IF






End Sub
Public Sub TreeKeyword_ondatachange()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Bitte zuerst das Ticket reservieren.")
  ELSE
    'Aktuellen Agent auslesen
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Datenbankverbindung zu helpline_replication
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Ditzingen oder TG auslesen
    Dim agentid, responsibility
    Set rs_resp = createobject("ADODB.Recordset")
    Set rs_resp = cn.Execute("Select responsibility from AgentID_responsibility where agentid = " & cstr(agent))
    responsibility = rs_resp.fields("responsibility").value
    rs_resp.close

    'Keyword einlesen
    Dim kw
    kw = hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 1)
    IF responsibility = 112545 THEN
      'KeywordOrga Wert aus Vergleichstabelle einlesen
      Dim kwo
      Set rs_kwkwo = createobject("ADODB.Recordset")
      Set rs_kwkwo = cn.Execute("Select keywordorga from kw_kwo_mapping where keywordid = " & cstr(kw))
      Do While Not rs_kwkwo.EOF
        kwo = rs_kwkwo.fields("keywordorga").value
        rs_kwkwo.MoveNext
      Loop
      IF Not kwo = "" THEN
        hlObj.SetValue "Keywords.KeywordOrga", 0, 0, 0, kwo
        TreeKeywordOrga.SelectTreeItem kwo
      END IF
      rs_kwkwo.close
    ELSE
      'Wert für die TG setzen
      'Dim tg
      'tg = HIER TG Value einlesen
      'hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
      'TreeKeywordOrga.SelectTreeItem tg
    END IF

    'Datenbankverbindung zu helpline_replication schließen
    cn.close
    Set cn = Nothing
  END IF

End Sub
Public Sub ComboLevel_SelectionChanged()
  'Bei Änderung des Supportlevels automatisch den Status auf "Weitergeleitet" setzen
  Dim level
  level = hlObj.GetValue("IncidentAttribute.EscalationLevel", 0, 0, 0, 0)

  IF level = "EscalationLevelLevel2" THEN
    hlObj.SetValue "IncidentAttribute.IncidentStatus", 0, 0, 0, "IncidentStatusRouted"
  END IF
  IF level = "EscalationLevelLevel1" THEN
    hlObj.SetValue "IncidentAttribute.IncidentStatus", 0, 0, 0, "IncidentStatusRouted"
  END IF

End Sub
Public Sub ButtonDiscovery_Click()
  Dim Hostname
  Hostname = hlProduct.getvalue("AssetGeneral.Hostname", 0, 0, 0, 0)
  Dim wshshell, oExec
  Set wshShell = CreateObject("Wscript.Shell")
  Command1 = "c:\program files\internet explorer\iexplore.exe http://srv01inv1/discovery/Reports/List.aspx?q=" + Hostname + "&flgDevice=1"
  Set oExec = wshShell.Exec(Command1)

End Sub
Public Sub b_template_save_Click()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Please reserve the ticket first.")

  ELSE

    'Templatenamen eingeben
    Dim name
    name = InputBox("Please type in a descriptive name for the template:", "templatename", "Maximum of 100 characters.")

    'Bei Abbruch nichts unternehmen, sonst weiter im Script
    IF name = FALSE THEN
    ELSE

      'Agentid auslesen anhand des aktuellen Agenten
      Dim agent
      agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

      'Datenbankverbindung zu helpline_replication
      Set cn = CreateObject("ADODB.Connection")
      'DB Verbindung öffnen
      cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
      cn.ConnectionTimeout = 10
      cn.Open

      'Teamname auslesen
      Dim teamID, teamDisplayname, agent_displayname
      Set rs_team = createobject("ADODB.Recordset")
      Set rs_team = cn.Execute("Select AgentTeam_ID,AgentTeam_Displayname,Agent_Displayname from IM_Agent_Supportteam where Agent_ID = " & cstr(agent))
      teamDisplayname = rs_team.fields("AgentTeam_Displayname").value
      teamID = rs_team.fields("AgentTeam_ID").value
      agent_displayname = rs_team.fields("Agent_Displayname").value
      rs_team.close

      'Abfrage ob Speicherung als persönliches oder als Teamtemplate gewünscht wird
      Dim result
      result = MsgBox("Button YES => personal template for: " & agent_displayname & chr(10) & chr(13) & chr(13) & "or" & chr(10) & chr(13) & chr(13) & "Button NO => team template for: ''" & teamDisplayname & "''", 4, "personal template or team template?")
      IF result = 6 THEN
        'Persönliches Insert auf Datenbank starten
        Set rs = cn.execute("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('" & cstr(agent) & "','" & name & "','" & hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0) & "','" & Replace(hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0), "'", "''") & "','" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, 0, 0), "'", "''") & "','" & Replace(hlObj.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0), "'", "''") & "','" & hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 0) & "','" & hlObj.GetValue("Keywords.KeywordOrga", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.EscalationLevel", 0, 0, 0, 0) & "','" & hlObj.GetValue("CaseClassificationAttribute.Impact", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.FunctionalRange", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0) & "','" & hlObj.GetValue("CaseGeneral.DefaultNotification", 0, 0, 0, 0) & "','" & cstr(agent) & "','" & hlObj.GetValue("IncidentAttribute.Convenience", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailTo", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailCC", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailSubject", 0, 0, 0, 0) & "')")
      ELSE
        'Team Insert auf Datenbank starten
        Set rs = cn.execute("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('" & cstr(teamID) & "','" & name & "','" & hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0) & "','" & Replace(hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0), "'", "''") & "','" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, 0, 0), "'", "''") & "','" & Replace(hlObj.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0), "'", "''") & "','" & hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 0) & "','" & hlObj.GetValue("Keywords.KeywordOrga", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.EscalationLevel", 0, 0, 0, 0) & "','" & hlObj.GetValue("CaseClassificationAttribute.Impact", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.FunctionalRange", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0) & "','" & hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0) & "','" & hlObj.GetValue("CaseGeneral.DefaultNotification", 0, 0, 0, 0) & "','" & cstr(agent) & "','" & hlObj.GetValue("IncidentAttribute.Convenience", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailTo", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailCC", 0, 0, 0, 0) & "','" & hlObj.GetValue("EmailSUAttribute.EmailSubject", 0, 0, 0, 0) & "')")

      END IF
      'Verbindung schließen
      cn.close

    END IF
  END IF


End Sub
Public Sub b_template_load_Click()
  'Prüfen ob Template in der Checkbox ausgewählt wurde
  IF cb_template_load.GetCurSel = - 1 or l_templateID.text = "" THEN
    Dim msg
    msg = MsgBox("Please select a template from the list." & Chr(13) & Chr(10) & "If the list is empty, there is no template existing.", vbOKOnly, "No data record available.")
  ELSE

    'Agentid auslesen anhand des aktuellen Agenten
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Angewählte ID aus Label auslesen
    Dim templateid
    templateid = l_templateID.Text

    'Datenbankverbindung zu helpline_replication
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Inhalte von agent_templates in das Recordset einlesen
    Set rs = createobject("ADODB.Recordset")
    Set rs = cn.Execute("Select * from templater where template_id = " & templateid)

    hlObj.SetValue "IncidentAttribute.RequestType", 0, 0, 0, rs.fields("Requesttype").value
    IF hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0) = "" THEN
      hlObj.SetValue "CaseDescription.DescriptionText", 0, 0, 0, rs.fields("descriptiontext").value
    ELSE
    END IF
    hlObj.SetValue "CaseDiagnosis.DiagnosisText", 0, 0, 0, rs.fields("diagnosistext").value
    hlObj.SetValue "CaseSolution.SolutionText", 0, 0, 0, rs.fields("solutiontext").value
    hlObj.SetValue "Keywords.Keyword", 0, 0, 0, rs.fields("keyword").value
    hlObj.SetValue "Keywords.KeywordOrga", 0, 0, 0, rs.fields("keywordorga").value
    hlObj.SetValue "IncidentAttribute.EscalationLevel", 0, 0, 0, rs.fields("EscalationLevel").value
    hlObj.SetValue "CaseClassificationAttribute.Impact", 0, 0, 0, rs.fields("Impact").value
    hlObj.SetValue "IncidentAttribute.FunctionalRange", 0, 0, 0, rs.fields("FunctionalRange").value
    hlObj.SetValue "IncidentAttribute.ProductionalRelevanz", 0, 0, 0, rs.fields("ProductionalRelevance").value
    hlObj.SetValue "EmailSUAttribute.EmailCaller", 0, 0, 0, rs.fields("EmailCaller").value
    hlObj.SetValue "IncidentAttribute.IncidentStatus", 0, 0, 0, rs.fields("IncidentStatus").value
    hlObj.SetValue "CaseGeneral.DefaultNotification", 0, 0, 0, rs.fields("DefaultNotification").value
    hlObj.SetValue "IncidentAttribute.Convenience", 0, 0, 0, rs.fields("PCAssoziated").value
    hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, rs.fields("EmailBodytext").value
    hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT", 0, 0, 0, rs.fields("EmailBodyRawtext").value
    'hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,rs.fields("EmailTo").value
    hlObj.SetValue "EmailSUAttribute.EmailCC", 0, 0, 0, rs.fields("EmailCC").value
    strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
    hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
    IF hlObj.GetValue("EmailSUAttribute.EmailSubject", 0, 0, 0, 0) = "" THEN
      hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, rs.fields("EmailSubject").value
    END IF

    'Subject Setzen
    varSubject = Left(EditProblem.Text, 100)
    IF EditSubjectCase.Text = "" THEN
      EditSubjectCase.Text = replace(varSubject, Chr(13) & Chr(10), " ")
    END IF

    'Übertrag der Caller in das An-Feld
    Dim tempmail
    tempmail = EditEmailAddress.text
    strEmail = ""
    CallerCount = 0
    CallerCount = hlObj.GetItemCount(0, 130)

    IF CallerCount > 0 THEN
      Dim CaseCallers
      Set CaseCallers = Nothing
      CaseCallers = hlObj.GetItems(0, - 1, - 1, 130)
      For Each Caller In CaseCallers
        CallerType = Caller.GetType
        IF CallerType = "Employee" THEN
          mailadr = ""
          mailadr = Caller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
          IF mailadr <> "" THEN
            strEmail = strEmail + mailadr + ";"
          END IF
        END IF
      Next
    ELSE
      strEmail = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
    END IF

    IF InStr(strEmail, tempmail) > 0 THEN
    ELSE
      strEmail = tempmail + ";" + strEmail
    END IF

    IF strEmail = "" THEN
      strEmail = hlObj.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
    END IF
    IF strEmail = "-" THEN
      strEmail = ""
    END IF

    'Aktivieren der Felder je nach EmailCaller Wert
    sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0)
    IF sendmail = "EmailCallerYes" THEN
      TextBoxEmailTo.Required = True
      TextBoxEmailSubject.Required = True
      GroupBoxEmail.Disabled = False
      hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
    ELSE
      TextBoxEmailTo.Required = False
      TextBoxEmailSubject.Required = False
      GroupBoxEmail.Disabled = True
    END IF

    'Aktivieren/Deaktivieren der Felder je nach gesetzter Anfrageart
    ComboProductionalRelevanz.Disabled = false
    IF Anfrageart <> "RequestTypeIncident" THEN
      ComboImpact.Disabled = True
      ComboFunctionalRange.Disabled = True
      hlObj.SetValue "CaseClassificationAttribute.Impact", 0, 0, 0, "ImpactOne"
      hlObj.SetValue "IncidentAttribute.FunctionalRange", 0, 0, 0, "FunctionalRangePartFailure"
      hlObj.SetValue "IncidentAttribute.ProductionalRelevanz", 0, 0, 0, "ProductionalRelevanzAdministrativeProcess"
    ELSE
      ComboImpact.Disabled = False
      ComboFunctionalRange.Disabled = False
      hlObj.SetValue "IncidentAttribute.ProductionalRelevanz", 0, 0, 0, "ProductionalRelevanzSupportProcess"
    END IF

    IF Anfrageart <> "RequestTypeContact" THEN
      CaseProblem.Disabled = False
      Dim status
      status = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
      IF Status <> "IncidentStatusClosed" THEN
        ComboBoxEmailCaller.Disabled = False
      ELSE
        ComboBoxEmailCaller.Disabled = True
      END IF
      CaseDiagnosis.Disabled = False
      KeywordTree.Disabled = False
      Attachment.Disabled = False
      CaseAttributes.Disabled = False
      ComboIncidentStatus.Disabled = False
    ELSE
      EditProblem.Text = ""
      CaseProblem.Disabled = True
      ComboBoxEmailCaller.Disabled = True
      EditDiagnosis.Text = ""
      CaseDiagnosis.Disabled = True
      KeywordTree.Disabled = True
      Attachment.Disabled = True
      CaseAttributes.Disabled = True
      ComboRequestType.Disabled = False
      ComboProductionalRelevanz.Disabled = true
      ComboIncidentStatus.Disabled = True
    END IF

    'Recordset schließen
    rs.close
    Set rs = Nothing

    'Datenbankverbindung zu helpline_replication schließen
    cn.close
    Set cn = Nothing

  END IF

End Sub
Public Sub b_template_change_Click()
  IF cb_template_load.GetCurSel = - 1 or l_templateID.text = "" THEN
    MsgBox("Please select template from list first.")
  ELSE

    'Agentid auslesen anhand des aktuellen Agenten
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Angewählte ID aus Label auslesen
    Dim templateid
    templateid = l_templateID.Text

    'Datenbankverbindung zu helpline_replication
    Set cn = CreateObject("ADODB.Connection")
    'DB Verbindung öffnen
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Recordset anlegen und templatenamen auslesen
    Dim templatename, editor
    Set rs = createobject("ADODB.Recordset")
    Set rs = cn.Execute("Select templatename,editor from templater where template_id = " & cstr(templateid))
    templatename = rs.fields("templatename").value
    editor = rs.fields("editor").value

    'Agent Name auslesen
    Dim agent_displayname
    Set rs_team = createobject("ADODB.Recordset")
    Set rs_team = cn.Execute("Select Agent_Displayname from IM_Agent_Supportteam where Agent_ID = " & cstr(editor))
    agent_displayname = rs_team.fields("Agent_Displayname").value
    rs_team.close

    'Nur wenn Agent = Editor überschreiben, sonst Abbruch
    IF editor <> cstr(agent) THEN
      Dim msg2
      msg2 = MsgBox("You can only overwrite self-created templates." & chr(10) & chr(13) & "template: " & templateid & " was created by: " & agent_displayname & "", vbOKOnly, "Overwrite is not allowed")
    ELSE
      Dim name
      name = InputBox("Please type in a descriptive name: ", "overwrite template", templatename)
      IF name = FALSE THEN
      ELSE

        'Abfrage ob Update erwünscht
        Dim result
        result = MsgBox("Möchten Sie das Template:  ''" & templatename & "''  überschreiben?", 4, "Template überschreiben?")
        IF result = 6 THEN

          'Update auf Datenbank wird ausgeführt
          Set rs = cn.execute("Update templater set templatename = '" & name & "', Requesttype = '" & hlObj.GetValue("IncidentAttribute.RequestType", 0, 0, 0, 0) & "',descriptiontext = '" & Replace(hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0), "'", "''") & "', diagnosistext = '" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, 0, 0), "'", "''") & "', solutiontext = '" & Replace(hlObj.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0), "'", "''") & "', keyword = '" & hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 0) & "', keywordorga = '" & hlObj.GetValue("Keywords.KeywordOrga", 0, 0, 0, 0) & "', EscalationLevel = '" & hlObj.GetValue("IncidentAttribute.EscalationLevel", 0, 0, 0, 0) & "',Impact = '" & hlObj.GetValue("CaseClassificationAttribute.Impact", 0, 0, 0, 0) & "',FunctionalRange = '" & hlObj.GetValue("IncidentAttribute.FunctionalRange", 0, 0, 0, 0) & "',ProductionalRelevance = '" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz", 0, 0, 0, 0) & "',EmailCaller = '" & hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0) & "',IncidentStatus = '" & hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0) & "',DefaultNotification = '" & hlObj.GetValue("CaseGeneral.DefaultNotification", 0, 0, 0, 0) & "',editor = '" & cstr(agent) & "',PCAssoziated = '" & hlObj.GetValue("IncidentAttribute.Convenience", 0, 0, 0, 0) & "',EmailBodyRawtext = '" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext", 0, 0, 0, 0) & "',EmailBodytext = '" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0) & "',EmailTo = '" & hlObj.GetValue("EmailSUAttribute.EmailTo", 0, 0, 0, 0) & "',EmailCC = '" & hlObj.GetValue("EmailSUAttribute.EmailCC", 0, 0, 0, 0) & "',EmailSubject = '" & hlObj.GetValue("EmailSUAttribute.EmailSubject", 0, 0, 0, 0) & "' where template_id = " & cstr(templateid))
          Set rs = nothing
        ELSE
        END IF

        'EndIF Überschreiben
      END IF

      'EndIf Agent = Editor
    END IF

    'Verbindung schließen
    cn.close

    'EndIf Wurde ein Checkbox-Wert zuvor angewählt
  END IF

End Sub
Public Sub cb_template_load_onfocus()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Please reserve the ticket first.")
    EditSurname.RequestFocus = true
  ELSE

    'Vorhandene Checkbox Werte entfernen
    cb_template_load.ResetContent

    'Agentid auslesen anhand des aktuellen Agenten
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Datenbankverbindung zu helpline_replication
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Teamname auslesen
    Dim teamID, teamDisplayname
    Set rs_team = createobject("ADODB.Recordset")
    Set rs_team = cn.Execute("Select AgentTeam_ID,AgentTeam_Displayname from IM_Agent_Supportteam where Agent_ID = " & cstr(agent))
    teamDisplayname = rs_team.fields("AgentTeam_Displayname").value
    teamID = rs_team.fields("AgentTeam_ID").value
    rs_team.close

    'Für Agent Templates ID bestimmen und selektierten Wert in Label schreiben
    Dim anzahl_agent_templates
    anzahl_agent_templates = 0
    Set rs = createobject("ADODB.Recordset")
    Set rs = cn.Execute("Select template_id,templatename from templater where agentid = " & cstr(agent) & " order by agentid, cast(Templatename as varchar(500))")
    ON ERROR RESUME NEXT
    rs.MoveFirst
    Do While Not rs.eof
      cb_template_load.AddItem(rs.fields("templatename").value)
      anzahl_agent_templates = anzahl_agent_templates + 1
      rs.MoveNext
    Loop

    'Trennlinie zwischen Agent-Templates einfügen
    cb_template_load.AddItem("---------------------------------Team templates below---------------------------------")

    'Für Team Templates ID bestimmen und selektierten Wert in Label schreiben
    Dim anzahl_team_templates
    anzahl_team_templates = 0
    Set rs2 = createobject("ADODB.Recordset")
    Set rs2 = cn.Execute("Select template_id,templatename from templater where agentid = " & cstr(teamID) & " order by agentid, cast(Templatename as varchar(500))")
    ON ERROR RESUME NEXT
    rs2.MoveFirst
    Do While Not rs2.eof
      cb_template_load.AddItem(rs2.fields("templatename").value)
      anzahl_team_templates = anzahl_team_templates + 1
      rs2.MoveNext
    Loop

    'Recordset schließen
    rs.close
    rs2.close


    'Datenbankverbindung zu helpline_replication schließen
    cn.close
    Set cn = Nothing

  END IF


End Sub
Public Sub b_template_delete_Click()
  'Prüfen ob Template in der Checkbox ausgewählt wurde
  IF cb_template_load.GetCurSel = - 1 or l_templateID.text = "" THEN
    Dim msg
    msg = MsgBox("Please select a template from the list." & Chr(13) & Chr(10) & "If the list is empty, there is no template existing.", vbOKOnly, "No data record available.")

  ELSE

    'Agentid auslesen anhand des aktuellen Agenten
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Angewählte ID aus Label auslesen
    Dim templateid
    templateid = l_templateID.Text

    'Datenbankverbindung zu helpline_replication
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Editor bestimmen
    Dim editor
    Set rs_editor = createobject("ADODB.Recordset")
    Set rs_editor = cn.Execute("Select editor from templater where template_id = " & cstr(templateid))
    editor = rs_editor.fields("editor").value
    rs_editor.close

    'Agent Name auslesen
    Dim agent_displayname
    Set rs_team = createobject("ADODB.Recordset")
    Set rs_team = cn.Execute("Select Agent_Displayname from IM_Agent_Supportteam where Agent_ID = " & cstr(editor))
    agent_displayname = rs_team.fields("Agent_Displayname").value
    rs_team.close

    IF editor <> cstr(agent) THEN
      Dim msg2
      msg2 = MsgBox("You are only allowed to delete self-created tickets." & chr(10) & chr(13) & "Template ID: " & templateid & " was created by:" & agent_displayname & "", vbOKOnly, "Delete not allowed.")
    ELSE

      'Abfrage ob Löschen erwünscht
      Dim result
      result = MsgBox("Do you really want to delete the template?", 4, "Delete template?")
      IF result = 6 THEN

        'Zeile von agent_templates löschen
        Set rs = createobject("ADODB.Recordset")
        Set rs = cn.Execute("Delete from templater where template_id = " & cstr(templateid))

        'Auswahl der Checkbox zurücksetzen und ID auf Null
        cb_template_load.ResetContent
        l_templateid.text = ""

        'Recordset schließen
        Set rs = Nothing
      ELSE
      END IF


      'End If Editor = Agent
    END IF

    'Datenbankverbindung zu helpline_replication schließen
    cn.close
    Set cn = Nothing

    'Vorhandene Checkbox Werte entfernen
    cb_template_load.ResetContent
    l_templateID.Text = ""

  END IF


End Sub
Public Sub cb_template_load_SelectionEndOK()
  'Agentid auslesen anhand des aktuellen Agenten
  Dim agent, team
  agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

  'Angewählte Position bestimmen
  Dim position
  position = cb_template_load.GetCurSel + 1

  'Datenbankverbindung zu helpline_replication
  Set cn = CreateObject("ADODB.Connection")
  cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
  cn.ConnectionTimeout = 10
  cn.Open

  'Teamname auslesen
  Dim teamID, teamDisplayname
  Set rs_teamid = createobject("ADODB.Recordset")
  Set rs_teamid = cn.Execute("Select AgentTeam_ID,AgentTeam_Displayname from IM_Agent_Supportteam where Agent_ID = " & cstr(agent))
  teamDisplayname = rs_teamid.fields("AgentTeam_Displayname").value
  teamID = rs_teamid.fields("AgentTeam_ID").value
  rs_teamid.close

  'Anzahl der Agenten-Templates bestimmen
  Dim anzahl_agent_templates
  anzahl_agent_templates = 0
  Set rs_anzahl = createobject("ADODB.Recordset")
  Set rs_anzahl = cn.Execute("Select template_id,templatename from templater where agentid = " & cstr(agent))
  ON ERROR RESUME NEXT
  rs_anzahl.MoveFirst
  Do While Not rs_anzahl.eof
    anzahl_agent_templates = anzahl_agent_templates + 1
    rs_anzahl.MoveNext
  Loop
  rs_anzahl.close

  IF position <= anzahl_agent_templates THEN
    'Select für Agententemplate ausführen
    Set rs_agent = createobject("ADODB.Recordset")
    Set rs_agent = cn.Execute("Select template_id from templater where agentid = '" & cstr(agent) & "' order by agentid, cast(Templatename as varchar(500))")
    ON ERROR RESUME NEXT
    rs_agent.MoveFirst
    For i = 1 To position
      l_templateID.Text = rs_agent.fields("template_id").value
      rs_agent.MoveNext
    Next
    'Dataset schließen
    rs_agent.close

  ELSE

    'Prüfung, ob Trennlinie ausgewählt wurde.
    IF cb_template_load.GetCurSel = anzahl_agent_templates THEN
      l_templateID.Text = ""
      'cb_template_load.ResetContent

    ELSE
      'Select für Teamtemplate ausführen  - "Position -1" wegen Trennzeile zwischen Templatetypen
      position = position - anzahl_agent_templates - 1
      Set rs_team = createobject("ADODB.Recordset")
      Set rs_team = cn.Execute("Select template_id from templater where agentid = '" & cstr(teamID) & "' order by agentid, cast(Templatename as varchar(500))")
      ON ERROR RESUME NEXT
      rs_team.MoveFirst
      For i = 1 To position
        l_templateID.Text = rs_team.fields("template_id").value
        rs_team.MoveNext
      Next
      'Dataset schließen
      rs_team.close

    END IF
  END IF

  'DB schließen
  cn.close

End Sub
Public Sub ButtonSCCMRemote_Click()
  Dim wshshell, oExec, OsType
  Set wshShell = CreateObject("Wscript.Shell")

  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)

  IF hlObj.IsReadOnly("CASEINFO.REACTIONTIME", 0) = 0 THEN

    objType = hlProduct.GetType
    IF objType = "DesktopComputer" Or objType = "ServerComputer" Or objType = "NotebookComputer" THEN
      'Auslesen des gewählten Computers
      'Reading the selected computer
      host = EditHostname.Text

      IF host <> "" THEN
        ON ERROR RESUME NEXT
        'Kommandozeile für den Aufruf von On Command Remote Master
        'Command lin for calling On Command Remote Master
        'Command1="""%programfiles%"\smsadmin\bin\i386\remote.exe 2 "" & host
        OsType = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
        IF OsType = 32 THEN
          'x86
          Command1 = "c:\Program Files\Microsoft Configuration Manager Console\AdminUI\bin\i386\rc.exe 1 " & host
        ELSE
          'x64
          Command1 = "c:\Program Files (x86)\Microsoft Configuration Manager Console\AdminUI\bin\i386\rc.exe 1 " & host
        END IF

        RemoteTool = "SCCM Remote"

        Set oExec = wshShell.Exec(Command1)
        IF err.Number = - 2147024893 THEN
          IF LangID = 7 THEN
            msgbox "Auf Ihrem Computer ist das Remote Tool " & RemoteTool & " nicht installiert." & vbLf & "Bitte wenden Sie sich an Ihren Administrator.", vbExclamation, "helpLine - ClassicDesk"
          ELSE
            msgbox "The remote tool " & RemoteTool & " is not installed on your computer." & vbLf & "Please consult your administrator.", vbExclamation, "helpLine - ClassicDesk"
          END IF
        END IF
      END IF
    ELSE
      IF LangID = 7 THEN
        msgbox "Es wurde kein Computer als Inventar ausgewählt." & vbLf & "Bitte wählen Sie einen Computer für den Vorgang aus.", vbExclamation, "helpLine - ClassicDesk"
      ELSE
        msgbox "No computer has been selected." & vbLf & "Please select a computer for this Case.", vbExclamation, "helpLine - ClassicDesk"
      END IF
    END IF
  END IF




End Sub
Public Sub ButtonShowOverView_Click()
  'Ermitteln der Locale ID für die Sprachauswahl
  'Selecting the Locale ID for the desired language
  lcid = hlSession.GetLocaleID
  LangID = hlSession.LangIDFromLCID(lcid)

  CaseOwner = hlObj.GetValue("HLOBJECTINFO.OWNER", 0, 0, 0, 0)
  Agent = ""
  IF LangID = 7 THEN
    Problemtitle = "<b>====== Problembeschreibung ======" & " [von Agent : " & CaseOwner & "]</b>" & vbNewLine
    Diagnosistitle = "<b>====== Kommunikation ======</b>" & vbNewLine
    Solutiontitle = "<b>====== Lösungsbeschreibung ======" & " [von Agent : " & hlObj.GetValue("SUINFO.EDITOR", 0, 0, 0, 0) & "]</b>" & vbNewLine
  ELSE
    Problemtitle = "<b>====== Problemdescription ======" & " [by Agent : " & CaseOwner & "]</b>" & vbNewLine
    Diagnosistitle = "<b>====== Diagnosisactivities ======</b>" & vbNewLine
    Solutiontitle = "<b>====== Final solution ======" & " [by Agent : " & hlObj.GetValue("SUINFO.EDITOR", 0, 0, 0, 0) & "]</b>" & vbNewLine
  END IF
  'VG-Beschreibung
  DescrText = ""
  DescrText = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 1, 0)
  IF DescrText = "" THEN
    DescrText = hlObj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0)
  END IF
  IF DescrText <> "" THEN
    DescrText = Replace(DescrText, vbCrLf, "<br>")
    ProblemAll = Problemtitle & DescrText & vbNewLine
  END IF
  'VG-Lösung
  'nur bei Status "Geschlossen" aus der aktuellen SU den Text holen
  actStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
  SolText = ""
  IF actStatus = "IncidentStatusClosed" THEN
    SolText = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0)
  END IF
  IF SolText = "" THEN
    SolText = hlObj.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0)
  END IF
  IF SolText <> "" THEN
    SolText = Replace(SolText, vbCrLf, "<br>")
    SolutionAll = Solutiontitle & SolText
  END IF

  SUIdx = hlObj.GetValue("SUINFO.INDEX", 0, 0, 0, 0)
  IF SUIdx > 0 THEN
    'Pro SU prüfen, ob Tätigkeitsbeschreibung eingetragen ist
    For i = 1 To SUIdx
      SUDiagnosisIntern = "<b> --- intern --- </b>"
      SUDiagnosis = ""
      SUDiagnosis = hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, i, 0)
      'SUDiagnosis = Replace(SUDiagnosis, Chr(13) & Chr(10), " ")
      SUDiagnosis = Replace(SUDiagnosis, vbCrLf, "<br>")
      SUDiagnosisExtern = "<b> --- extern --- </b>"
      SUDiagnosisExt = ""
      SUDiagnosisExt = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, i, 0)
      IF SUDiagnosis <> "" THEN
        SUActivity = hlObj.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, i, 0)
        SURegTime = hlObj.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, i, 0)
        Agent = hlObj.GetValue("SUINFO.EDITOR", 0, 0, i, 0)
        DiagnosisAll = DiagnosisAll & SUDiagnosisIntern & vbNewLine & "<b>" & i & ". SU (" & Agent & ") -> " & SUActivity & " [" & SURegTime & "]:" & "</b>" & vbNewLine & SUDiagnosis & vbNewLine & String(80, "-") & vbNewLine
      END IF
      IF SUDiagnosisExt <> "" THEN
        'SUDiagnosisExt = Replace(SUDiagnosisExt, vbCrLf, "<br>")
        SUActivity = hlObj.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, i, 0)
        SURegTime = hlObj.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, i, 0)
        Agent = hlObj.GetValue("SUINFO.EDITOR", 0, 0, i, 0)
        DiagnosisAll = DiagnosisAll & SUDiagnosisExtern & vbNewLine & "<b>" & i & ". SU (" & Agent & ") -> " & SUActivity & " [" & SURegTime & "]:" & "</b>" & vbNewLine & SUDiagnosisExt & vbNewLine & String(80, "-") & vbNewLine
      END IF
    Next
  END IF
  IF DiagnosisAll <> "" THEN
    DiagnosisAll = Diagnosistitle & DiagnosisAll
  END IF
  ProblemAll = ProblemAll & DiagnosisAll & SolutionAll
  'hlObj.SetValue "CaseGeneral.Overview",0,0,0,ProblemAll
  ProblemAll = Replace(ProblemAll, vbCrLf, "<br>")
  hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE", 0, 0, 0, ProblemAll

  'Button nach 1. Klick sperren
  'ButtonShowOverView.Disabled = True

End Sub
Public Sub ComboBoxEmailCaller_SelectionChanged()
  sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller", 0, 0, 0, 0)
  strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
  Dim tempmail
  tempmail = EditEmailAddress.text
  'Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
  'Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
  Dim strIncStatus
  strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
  strSubject = hlObj.GetValue("CaseGeneral.Subject", 0, 0, 0, 0)
  strEmail = ""
  CallerCount = 0
  CallerCount = hlObj.GetItemCount(0, 130)

  IF CallerCount > 0 THEN
    Dim CaseCallers
    Set CaseCallers = Nothing
    CaseCallers = hlObj.GetItems(0, - 1, - 1, 130)
    For Each Caller In CaseCallers
      CallerType = Caller.GetType
      IF CallerType = "Employee" THEN
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
        IF mailadr <> "" THEN
          strEmail = strEmail + mailadr + ";"
        END IF
      END IF
    Next

  ELSE
    strEmail = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF

  IF InStr(strEmail, tempmail) > 0 THEN
  ELSE
    strEmail = tempmail + ";" + strEmail
  END IF

  IF strEmail = "" THEN
    strEmail = hlObj.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF
  IF strEmail = "-" THEN
    strEmail = ""
  END IF
  IF sendmail = "EmailCallerYes" THEN
    hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail
    hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, strSubject
    TextBoxEmailTo.Required = True
    TextBoxEmailSubject.Required = True
    GroupBoxEmail.Disabled = False
  ELSE
    hlObj.SetValue "EmailSUAttribute.EmailSearchName", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailSearchResult", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailCC", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailSubject", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, ""
    hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT", 0, 0, 0, ""
    GroupBoxEmail.Disabled = True
    TextBoxEmailTo.Required = False
    TextBoxEmailSubject.Required = False
  END IF

End Sub
Public Sub ButtonSearchMail_Click()
  'EMail-Adressen leeren
  ComboBoxEmailSearchResult.Text = ""
  ComboBoxEmailSearchResult.ResetContent
  'Name als Suchparameter für Email-Adressen abfragen
  Name = TextBoxEmailSearchName.Text

  Dim ConString
  'ConString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm4t"
  ConString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm1"


  'Dim i
  'i = 0

  IF Name <> "" THEN
    '------------------------------------------------------------------------------------------------
    'Ermitteln der Email-Adressen auf Bases des eingegebenen Namens
    Set cn2 = createobject("ADODB.Connection")

    'Verbindung öffnen
    cn2.ConnectionString = ConString
    cn2.ConnectionTimeout = 10
    cn2.Open

    'SELECT absetzen
    Set rs2 = createobject("ADODB.Recordset")
    Set rs2 = cn2.Execute("select email from _EMails where email LIKE '%" & Name & "%' ORDER BY email")

    'Daten einlesen
    Data = ""
    Do While Not rs2.eof
      'In Variable schreiben
      i = i + 1
      ComboBoxEmailSearchResult.AddItem rs2.fields(0).value
      IF i = 1 THEN
        ComboBoxEmailSearchResult.Text = rs2.fields(0).value
      END IF
      rs2.movenext
    Loop
    'Verbindung schließen
    rs2.close
    cn2.close

  END IF






End Sub
Public Sub ButtonTo_Click()
  email = ComboBoxEmailSearchResult.Text
  Recipient = TextBoxEmailTo.Text
  IF email = "" THEN
    MsgBox "Bitte eine Email-Adresse auswählen!"
  ELSE
    fullemailstring = len(email)
    pos = Instr(1, email, ":", 1)
    emailstring = clng(fullemailstring) - clng(pos)
    email = Right(email, CLNG(emailstring))
    IF Recipient = "" THEN
      Recipient = email
    ELSE
      IF RIGHT(Recipient, 1) = ";" THEN
        Recipient = Recipient + email
      ELSE
        Recipient = Recipient + ";" + email
      END IF
    END IF
    TextBoxEmailTo.Text = Recipient
  END IF

End Sub
Public Sub ButtonCC_Click()
  email = ComboBoxEmailSearchResult.Text
  RecipientCC = TextBoxEmailCC.Text
  IF email = "" THEN
    MsgBox "Bitte eine Email-Adresse auswählen!"
  ELSE
    fullemailstring = len(email)
    pos = Instr(1, email, ":", 1)
    emailstring = clng(fullemailstring) - clng(pos)
    email = Right(email, CLNG(emailstring))
    IF RecipientCC = "" THEN
      RecipientCC = email
    ELSE
      IF RIGHT(RecipientCC, 1) = ";" THEN
        RecipientCC = RecipientCC + email
      ELSE
        RecipientCC = RecipientCC + ";" + email
      END IF
    END IF
    TextBoxEmailCC.Text = RecipientCC
  END IF


End Sub
Public Sub ButtonSetAgent1_Click()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Bitte zuerst das Ticket reservieren.")
  ELSE

    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Datenbankverbindung zu helpline_data
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn.ConnectionTimeout = 10
    cn.Open

    'Teamname auslesen
    Dim agentid, internalname
    Set rs_kwo = createobject("ADODB.Recordset")
    Set rs_kwo = cn.Execute("Select name,internalname from vw_agent_to_first_keywordorga where agentid = " & cstr(agent))
    internalname = rs_kwo.fields("internalname").value

    'Wert in Schlagwort schreiben
    hlObj.SetValue "Keywords.KeywordOrga", 0, 0, 0, internalname
    TreeKeywordOrga.SelectTreeItem internalname

    'Datenbankverbindung zu helpline_replication schließen
    rs_kwo.close
    cn.close
    Set cn = Nothing

  END IF





End Sub
Public Sub ButtonSetKW_Click()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Bitte zuerst das Ticket reservieren.")
  ELSE
    'Aktuellen Agent auslesen
    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Datenbankverbindung zu helpline_replication
    Set cn1 = CreateObject("ADODB.Connection")
    cn1.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn1.ConnectionTimeout = 10
    cn1.Open

    'Teamname auslesen
    Dim keywordid
    Set rs_kw = createobject("ADODB.Recordset")
    Set rs_kw = cn1.Execute("Select keywordid from vw_Agent_Emplkeyword where agentid = " & cstr(agent))
    keywordid = rs_kw.fields("keywordid").value
    rs_kw.close

    'Wert in Schlagwort schreiben
    hlObj.SetValue "Keywords.Keyword", 0, 0, 0, keywordid
    TreeKeyword.SelectTreeItem keywordid
    TreeKeyword.ExpandTreeItem keywordid

    'Responsibility - Ditzingen oder TG - einlesen
    Dim responsibility
    Set rs_resp = createobject("ADODB.Recordset")
    Set rs_resp = cn1.Execute("Select responsibility from AgentID_responsibility where agentid = " & cstr(agent))
    responsibility = rs_resp.fields("responsibility").value
    rs_resp.close

    'Keyword einlesen
    Dim kw
    kw = hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 1)
    IF responsibility = 112545 THEN
      'KeywordOrga Wert aus Vergleichstabelle einlesen
      Dim kwo
      Set rs_kwkwo = createobject("ADODB.Recordset")
      Set rs_kwkwo = cn1.Execute("Select keywordorga from kw_kwo_mapping where keywordid = " & cstr(kw))
      Do While Not rs_kwkwo.EOF
        kwo = rs_kwkwo.fields("keywordorga").value
        rs_kwkwo.MoveNext
      Loop
      IF Not kwo = "" THEN
        hlObj.SetValue "Keywords.KeywordOrga", 0, 0, 0, kwo
        TreeKeywordOrga.SelectTreeItem kwo
      END IF
      rs_kwkwo.close
    ELSE
      'Wert für die TG setzen
      'Dim tg
      'tg = HIER TG Value einlesen
      'hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
      'TreeKeywordOrga.SelectTreeItem tg
    END IF

    'Datenbankverbindung zu helpline_replication schließen
    cn1.close
    Set cn1 = Nothing
  END IF





End Sub
Public Sub ButtonResetTo_Click()
  CallerCount = 0
  CallerCount = hlObj.GetItemCount(0, 130)
  IF CallerCount > 0 THEN
    Dim CaseCallers
    Set CaseCallers = Nothing
    CaseCallers = hlObj.GetItems(0, - 1, - 1, 130)
    For Each Caller In CaseCallers
      CallerType = Caller.GetType
      IF CallerType = "Employee" THEN
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
        IF mailadr <> "" THEN
          strEmail = strEmail + mailadr + ";"
        END IF
      END IF
    Next
  ELSE
    strEmail = hlCaller.GetValue("PersonInformation.EmailAddress", 0, 0, 0, 0)
  END IF

  Dim tempmail
  tempmail = EditEmailAddress.text
  IF InStr(strEmail, tempmail) > 0 THEN
  ELSE
    strEmail = tempmail + ";" + strEmail
  END IF

  hlObj.SetValue "EmailSUAttribute.EmailTo", 0, 0, 0, strEmail

End Sub
Public Sub ButtonEmailPreview_Click()
  status = hlObj.GetValue("IncidentAttribute.IncidentStatus", 0, 0, 0, 0)
  HLinkToCase = "http://srv01itsm2/helpLinePortal"
  HTicketID = hlobj.GetValue("CASEINFO.REFERENCENUMBER", 0, 0, 0, 0)
  SubjectCase = hlobj.GetValue("EmailSUAttribute.EmailSubject", 0, 0, 0, 0)
  LanguageDE = 0
  MailTo = hlobj.GetValue("EmailSUAttribute.EmailTo", 0, 0, 0, 0)
  For z = 1 To len(MailTo)
    IF Mid(MailTo, z, 1) = "@" THEN
      CounterEmpf = CounterEmpf + 1
    END IF
  Next
  IF IsObject(hlCaller) = True THEN
    surname = hlCaller.GetValue("PersonGeneral.PersonSurname", 0, 0, 0, 0)
    letteraddress = hlCaller.GetValue("PersonGeneral.ShortLetterAddress", 0, 0, 0, 0)
    language = hlCaller.GetValue("PersonGeneral.Language", 0, 0, 0, 0)
    IF language <> "LanguageGerman" THEN
      LanguageDE = - 1
    ELSE
      LanguageDE = 1
    END IF
  ELSE
    surname = "Unbekannt/Unknown"
  END IF
  Editor = hlobj.GetValue("SUINFO.EDITOR", 0, 0, 0, 0)
  '----------------------------------------------------------------------------------------------------------
  'M.Rettig, 14.05.2012 - SU-Email als HTML-Vorschau
  IF status = "IncidentStatusClosed" THEN
    Const ForReading = 1, ForWriting = 2, ForAppending = 8

    Dim OriginDescr
    OriginDescr = hlobj.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0)
    MailBody = hlobj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0)
    'Deutsche Werte
    IF LanguageDE > 0 THEN
      IF letteraddress = "" THEN
        letteraddress = "Herr/Frau"
      END IF

      'Konstante Werte deutsch setzen
      TTicketID = "Ticketnummer"
      TStatus = "Status"
      HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus", 7, 0, LastSUIdx, 0)
      TEditor = "Bearbeiter"
      TSubject = "Betreff:"
      IF CounterEmpf > 1 THEN
        Anrede = "Sehr geehrte "
        surname = "Damen und Herren"
      ELSE
        Anrede = "Sehr geehrte(r) " & CStr(letteraddress)
      END IF
      TSolution = "Lösung:"
      TBeschr = "Ticket-Beschreibung:"
      TComplimentary = "Mit freundlichen Grüßen,"
      TSignature = "Ihr Team IT + Prozesse"
      TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!"
      Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME", 7, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      subject = "Lösung zur IT Service Desk Anfrage " & " [#"
      subject = subject & HTicketID & "]" & " vom " & Datum
      TIntroduction = "Wir möchten Ihnen folgende Lösung übermitteln:"
    ELSE
      IF letteraddress = "" THEN
        letteraddress = "Mrs./Ms./Mr."
      END IF

      'Konstante Werte englisch setzen
      TTicketID = "Ticket number"
      TStatus = "Status"
      HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus", 9, 0, LastSUIdx, 0)
      TEditor = "Editor"
      TSubject = "Subject:"
      IF CounterEmpf > 1 THEN
        Anrede = "Dear "
        surname = "Sir or Madam"
      ELSE
        Anrede = "Dear " & CStr(letteraddress)
      END IF
      TSolution = "Solution:"
      TBeschr = "Ticket-Description:"
      TComplimentary = "Best regards,"
      TSignature = "Your support team IT + Processes"
      TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!"
      Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME", 9, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      subject = "Your support request from " & Datum & " with the reference no. [#"
      subject = subject & HTicketID & "]"
      TIntroduction = "We deliver to you the following solution description:"
    END IF
    MailBody = Replace(MailBody, vbCrLf, "<br>")
    OriginDescr = Replace(OriginDescr, vbCrLf, "<br>")
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Öffnet das File zum lesen
    Set f = fso.OpenTextFile("C:\TRUMPF\helpline\Emailtemplate.html", ForReading)
    'Liest alle Daten in die Variable BodyText
    BodyText = f.ReadAll
    BodyText = replace(BodyText, "[$NoticeTop$]", TNoticeTop)
    BodyText = replace(BodyText, "[$Ticket-ID_Titel$]", TTicketID)
    BodyText = replace(BodyText, "[$TicketID$]", HTicketID)
    BodyText = replace(BodyText, "[$Ticketstatus_Titel$]", TStatus)
    BodyText = replace(BodyText, "[$Ticketstatus$]", HStatus)
    BodyText = replace(BodyText, "[$Editor_Titel$]", TEditor)
    BodyText = replace(BodyText, "[$Editor$]", Editor)
    BodyText = replace(BodyText, "[$CaseSubject_Titel$]", TSubject)
    BodyText = replace(BodyText, "[$CaseSubject$]", SubjectCase)
    BodyText = replace(BodyText, "[$LinktoCase_Titel$]", HLinkToCase)
    BodyText = replace(BodyText, "[$Salutation$]", Anrede)
    BodyText = replace(BodyText, "[$LastnameUser$]", cstr(surname))
    BodyText = replace(BodyText, "[$Introduction$]", TIntroduction)
    BodyText = replace(BodyText, "[$CaseSolution_Titel$]", TSolution)
    BodyText = replace(BodyText, "[$CaseSolution$]", MailBody)
    BodyText = replace(BodyText, "[$CaseDescription_Titel$]", TBeschr)
    BodyText = replace(BodyText, "[$CaseDescription$]", OriginDescr)
    BodyText = replace(BodyText, "[$ComplimentaryClose$]", TComplimentary)
    BodyText = replace(BodyText, "[$Signature$]", TSignature)
    BodyText = replace(BodyText, "[$NoticeBottom$]", TNoticeBottom)
    'Schließt das File
    f.Close
    Set f = Nothing
    Set fso = Nothing
    hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT", 0, 0, 0, BodyText
    hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE", 0, 0, 0, BodyText
  ELSE
    DiagnText = hlobj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, 0, 0)
    IF LanguageDE = 1 THEN
      IF letteraddress = "" THEN
        letteraddress = "Herr/Frau"
      END IF

      'Konstante Werte deutsch setzen
      TTicketID = "Ticketnummer"
      TStatus = "Status"
      HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus", 7, 0, LastSUIdx, 0)
      TEditor = "Bearbeiter"
      TSubject = "Betreff:"
      IF CounterEmpf > 1 THEN
        Anrede = "Sehr geehrte "
        surname = "Damen und Herren"
      ELSE
        Anrede = "Sehr geehrte(r) " & CStr(letteraddress)
      END IF
      TDiagnosis = "Zwischenbescheid"
      TResubTime = "Wiedervorlagedatum:"
      TComplimentary = "Mit freundlichen Grüßen,"
      TSignature = "Ihr Team IT + Prozesse"
      TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!"

      'Hier wird die Betreffzeile erstellt
      'The subject field is entered here
      Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME", 7, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      ResubmissionTime = hlobj.GetValue("CASEINFO.RESUBMISSIONTIME", 7, 0, 0, 0)
      IF ResubmissionTime <> "" THEN
        IF DateDiff("d", Now, ResubmissionTime) > 0 THEN
          'If ResubmissionTime > Now Then
          ResubDatum = MID(ResubmissionTime, 1, 10)
        ELSE
          ResubDatum = ""
        END IF
      END IF
      subject = "Zwischenbescheid zur IT Service Desk Anfrage " & " [#"
      subject = subject & HTicketID & "]" & " vom " & Datum
      TIntroduction = "Wir möchten Ihnen folgende Nachricht übermitteln:"
    ELSE
      IF letteraddress = "" THEN
        letteraddress = "Mrs./Ms./Mr."
      END IF

      'Konstante Werte englisch setzen
      TTicketID = "Ticket number"
      TStatus = "Status"
      HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus", 9, 0, LastSUIdx, 0)
      TEditor = "Editor"
      TSubject = "Subject:"
      IF CounterEmpf > 1 THEN
        Anrede = "Dear "
        surname = "Sir or Madam"
      ELSE
        Anrede = "Dear " & CStr(letteraddress)
      END IF

      TDiagnosis = "Intermediate Reply"
      TResubTime = "Resubmissiontime:"
      TComplimentary = "Best regards,"
      TSignature = "Your support team IT + Processes"
      TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!"


      'Hier wird die Betreffzeile erstellt
      'The subject field is entered here
      Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME", 9, 0, 0, 0)
      Datum = Mid(Creationdate, 1, 10)
      ResubmissionTime = hlobj.GetValue("CASEINFO.RESUBMISSIONTIME", 9, 0, 0, 0)
      IF ResubmissionTime <> "" THEN
        IF DateDiff("d", Now, ResubmissionTime) > 0 THEN
          'If ResubmissionTime > Now Then
          ResubDatum = MID(ResubmissionTime, 1, 10)
        ELSE
          ResubDatum = ""
        END IF
      END IF
      subject = "Your support request from " & Datum & " with the reference no. [#"
      subject = subject & HTicketID & "]"
      TIntroduction = "We deliver to you the following processing description:"
    END IF

    'Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Öffnet das File zum lesen
    Set f = fso.OpenTextFile("C:\TRUMPF\helpLine\IntermediateReply.html", ForReading)
    'Liest alle Daten in die Variable BodyText
    BodyText = f.ReadAll
    BodyText = replace(BodyText, "[$NoticeTop$]", TNoticeTop)
    BodyText = replace(BodyText, "[$Ticket-ID_Titel$]", TTicketID)
    BodyText = replace(BodyText, "[$TicketID$]", HTicketID)
    BodyText = replace(BodyText, "[$Ticketstatus_Titel$]", TStatus)
    BodyText = replace(BodyText, "[$Ticketstatus$]", HStatus)
    BodyText = replace(BodyText, "[$Editor_Titel$]", TEditor)
    BodyText = replace(BodyText, "[$Editor$]", Editor)
    BodyText = replace(BodyText, "[$CaseSubject_Titel$]", TSubject)
    BodyText = replace(BodyText, "[$CaseSubject$]", SubjectCase)
    BodyText = replace(BodyText, "[$LinktoCase_Titel$]", HLinkToCase)
    BodyText = replace(BodyText, "[$Salutation$]", Anrede)
    BodyText = replace(BodyText, "[$LastnameUser$]", cstr(surname))
    BodyText = replace(BodyText, "[$Introduction$]", TIntroduction)
    BodyText = replace(BodyText, "[$CaseInformation_Titel$]", TDiagnosis)
    BodyText = replace(BodyText, "[$CaseInformation$]", DiagnText)
    IF ResubDatum <> "" THEN
      BodyText = replace(BodyText, "[$ResubmissionTime_Titel$]", TResubTime)
      BodyText = replace(BodyText, "[$ResubmissionTime$]", ResubDatum)
    ELSE
      BodyText = replace(BodyText, "[$ResubmissionTime_Titel$]", "")
      BodyText = replace(BodyText, "[$ResubmissionTime$]", "")
    END IF
    BodyText = replace(BodyText, "[$ComplimentaryClose$]", TComplimentary)
    BodyText = replace(BodyText, "[$Signature$]", TSignature)
    'Schließt das File
    f.Close
    Set f = Nothing
    Set fso = Nothing
    hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT", 0, 0, 0, BodyText
    hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE", 0, 0, 0, BodyText
  END IF


End Sub
Public Sub ButtonSaveKW_Click()
  Dim isreserved
  isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 0)
  IF isreserved = "" THEN
    MsgBox("Bitte zuerst das Ticket reservieren.")
  ELSE

    Dim agent
    agent = hlObj.GetValue("CASEINFO.RESERVEDBY", 0, 0, 0, 1)

    'Datenbankverbindung zu helpline_replication
    Set cn1 = CreateObject("ADODB.Connection")
    cn1.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
    cn1.ConnectionTimeout = 10
    cn1.Open

    'Keyword einlesen und in Datenbank ablegen
    Dim personid, keywordid
    keywordid = hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 1)
    IF (CDbl(keywordid)) > 0 THEN
      'Personid über AgentID ermitteln
      Set rs_person = createobject("ADODB.Recordset")
      Set rs_person = cn1.Execute("Select personid from vw_Agent_Emplkeyword where agentid = " & cstr(agent))
      personid = rs_person.fields("personid").value
      rs_person.close

      'Datenbankverbindung zu helpline_data
      Set cn = CreateObject("ADODB.Connection")
      cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2"
      cn.ConnectionTimeout = 10
      cn.Open
      'Keyword schreiben
      Set rs_kw = createobject("ADODB.Recordset")
      Set rs_kw = cn.Execute("Update dbo.emplkeywords set keyword = " & cdbl(hlObj.GetValue("Keywords.Keyword", 0, 0, 0, 1)) & " where personid = " & cstr(personid))
      'Datenbank schließen
      'rs_kw.close
      cn.close
      Set cn = Nothing
    ELSE
      MsgBox("Please select a keyword first.")
    END IF


    'Datenbankverbindung zu helpline_replication schließen
    cn1.close
    Set cn1 = Nothing

  END IF






End Sub
Public Sub EditSubjectCase_ondatachange()
  Dim Text
  IF InStr(1, EditSubjectCase.Text, "Notfalltransport_SAP", vbTextCompare) THEN
    CaseProblem.Disabled = False
    CaseProblem.Disabled = False
    ComboBoxEmailCaller.Disabled = False
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    CaseAttributes.Disabled = False
    ComboIncidentStatus.Disabled = False
  END IF

  IF InStr(1, EditSubjectCase.Text, "Systemänderbarkeit_SAP", vbTextCompare) THEN
    CaseProblem.Disabled = False
    CaseProblem.Disabled = False
    ComboBoxEmailCaller.Disabled = False
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    CaseAttributes.Disabled = False
    ComboIncidentStatus.Disabled = False
  END IF

  IF InStr(1, EditSubjectCase.Text, "#Prio 1 Incident# ", vbTextCompare) THEN
    CaseProblem.Disabled = False
    CaseProblem.Disabled = False
    ComboBoxEmailCaller.Disabled = False
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    CaseAttributes.Disabled = False
    ComboIncidentStatus.Disabled = False
  END IF
  IF InStr(1, EditSubjectCase.Text, "Debugg_Modus_SAP", vbTextCompare) THEN
    CaseProblem.Disabled = False
    CaseProblem.Disabled = False
    ComboBoxEmailCaller.Disabled = False
    CaseDiagnosis.Disabled = False
    KeywordTree.Disabled = False
    Attachment.Disabled = False
    CaseAttributes.Disabled = False
    ComboIncidentStatus.Disabled = False
  END IF

End Sub
Public Sub ButtonActionItemsAdd_Click()
  Dim textdata, texttemp
  IF TextBoxActionItemsInput.Text = "" THEN
    MsgBox("Input value is missing.")
  ELSE
    texttemp = TextBoxActionItemsInput.Text
    textdata = hlObj.GetValue("IncidentAttribute.ActionItems", 0, 0, 0, 0)
    IF Not textdata = "" THEN
      textdata = textdata & CHR(10) & texttemp
    ELSE
      textdata = texttemp
    END IF
    hlObj.SetValue "IncidentAttribute.ActionItems", 0, 0, 0, textdata
  END IF

End Sub
Public Sub ButtonActionItemsDel_Click()
  Dim delete
  delete = MsgBox("Delete all action items permanently?", 4, "Delete Action Items")
  IF delete = 6 THEN
    hlObj.SetValue "IncidentAttribute.ActionItems", 0, 0, 0, ""
  END IF

End Sub
