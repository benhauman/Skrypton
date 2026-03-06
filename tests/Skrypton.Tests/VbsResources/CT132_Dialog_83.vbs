SUB IncReqOnLoad()
        Dim ReadOnly, NoPerson, NoAsset
        ReadOnly = True
        NoPerson = True
        NoAsset = True

        'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
        'First of all check whether the Case is write protected
        If hlObj.IsReadOnly("CaseGeneral.Subject",0)=0 Then
        ReadOnly=False
        End If

        'Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
        'Check wether the Caller object exist
        If IsObject(hlCaller) = True And EditSurname.Text <> "" Then
        NoPerson = False
        End If

        'VIP-Status des Anfragers abfragen und im Vorgang setzen
        Valid = hlCaller.HasContent("PersonGeneral.VIPLevel",0,0)
        If Valid = 1 Then
        VIP = hlCaller.GetValue("PersonGeneral.VIPLevel",0,0,0,0)
        'If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
        Select Case vip
        Case "VIPLevelVIP"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,1
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(142,139,254)
        Case "VIPLevelITAdminDitzingen"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,2
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelITAdminTG"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,3
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelSAPKeyUserTUS"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,4
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelNon"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,0
        ComboVIPStatus.Disabled = True
        Person.BackColor = ""
        End Select
        End If

        'Prüft ob ein Produkt Objekt vorhanden ist und ob dieses auch angezeigt wird
        'Check wether the Product object exist
        If IsObject(hlProduct) = True And EditAssetModel.Text <> "" Then
        NoAsset = False
        End If

        'Ermitteln der Locale ID für die Sprachauswahl
        'Selecting the Locale ID for the desired language
        lcid = hlSession.GetLocaleID
        LangID = hlSession.LangIDFromLCID(lcid)

        'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check requester search status to set the caption of the button
        If NoPerson = False Then
        If SearchCaller.GetSearchState = 3 Then
        SearchCaller.Caption = "Reset"
        Else
        SearchCaller.Caption = "Betroffener"
        End If
        End If

        'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check Asset search status to set the caption of the button
        If NoAsset = False Then
        If SearchAsset.GetSearchState = 3 Then
        SearchAsset.Caption = "Reset"
        Else
        SearchAsset.Caption = "Inventar"
        End If
        End If

        If NoAsset = False Then
        'Setzen des Inventars
        'Setting the asset
        varString=""
        varAType = hlProduct.GetType ()
        If varAType = "DesktopComputer" Or varAType = "ServerComputer" Or varAType = "NotebookComputer" Or varAType = "Printer" Then
        If EditHostname.Text <> "" Then varString = EditHostname.Text
        If EditAssetModel.Text<>"" Then varString = varString & " " & EditAssetModel.Text
        Else
        If EditAssetModel.Text<>"" Then varString=EditAssetModel.Text Else EditAssetModel.Text=" "
        End If
        EditAssetModel.Text = varString
        End If

        'Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
        Dim Anfrageart
        Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0)

        If Anfrageart <> "RequestTypeIncident" Then
        ComboImpact.Disabled = True
        ComboFunctionalRange.Disabled = True
        Else
        ComboImpact.Disabled = False
        ComboFunctionalRange.Disabled = False
        End If

        If Anfrageart <> "RequestTypeContact" Then
        CaseProblem.Disabled = False
        ComboBoxEmailCaller.Disabled = False
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        ComboIncidentStatus.Disabled = False

        Else
        CaseProblem.Disabled = True
        ComboBoxEmailCaller.Disabled = True
        CaseDiagnosis.Disabled = True
        KeywordTree.Disabled = True
        Attachment.Disabled = True
        ComboProductionalRelevanz.Disabled = true
        ComboIncidentStatus.Disabled = True
        End If

        'Zugriff auf Übersichts-Buttons regeln
        If ReadOnly = False Then
        ButtonShowOverView.Disabled = False
        ButtonEmailPreview.Disabled = False
        EditSubjectCase.Disabled = False
        Else
        ButtonShowOverView.Disabled = True
        ButtonEmailPreview.Disabled = True
        EditSubjectCase.Disabled = True
        End if

        'Einfärben der GrupBox CaseAttributes je nach Priorität
        Select Case hlObj.GetValue("CaseClassificationAttribute.Priority",0,0,0,0)
        Case "Priority1"
        CaseAttributes.BackColor = RGB(107,105,248)
        Case "Priority2"
        CaseAttributes.BackColor = RGB(119,170,251)
        Case "Priority3"
        CaseAttributes.BackColor = RGB(132,235,255)
        Case "Priority4"
        CaseAttributes.BackColor = RGB(128,213,177)
        Case "Priority5"
        CaseAttributes.BackColor = RGB(123,190,99)
        Case Else
        CaseAttributes.BackColor = RGB(248,245,240)
        End Select

        'Bei Status ToProof wird die Email-Tab angewählt
        If hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0) = "IncidentStatusToProof" then
        TabPageEmail.UiActive = True
        Else
        End If
      
END SUB
SUB OnSUIDAdded()
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
        If hlObj.IsReadOnly("CaseGeneral.Subject",0)=0 Then
        ReadOnly=False
        End If


        'Status auf "In Bearbeitung" setzen
        hlObj.SetValue"IncidentAttribute.IncidentStatus",0,0,0,"IncidentStatusInProgress"

        'Wenn Vorgang erweitert wird, wird die Zuständigkeit des Agenten ermittelt und gestezt.
        Dim GetLastSUIdx
        GetLastSUIdx = 0
        Dim suindices
        suindices = hlobj.GetSvcUnitIndices()
        GetLastSUIdx = UBound(suindices)
        If GetLastSUIdx > 0 Then
        Dim agent
        agent = hlObj.GetValue ("SUINFO.EDITOR",0,0,GetLastSUIdx+1,1)
        Dim person, helper, responsibilty
        Set helper = CreateObject("helpline.hlcontrols.HLHelperPFA")
        Set person = helper.GetPersonForAgent(model.GetClientContext,clng(agent))
        If isObject(person) = True Then
        responsibility = person.GetValue("PersonGeneralTrumpf.Responsibility",0,0,0,0)
        If responsibility = "ResponsibilityBSZDitzingen" Then
        hlObj.SetValue "IncidentAttribute.Responsibility",0,0,0,"ResponsibilityBSZDitzingen"
        Else
        hlObj.SetValue "IncidentAttribute.Responsibility",0,0,0,"ResponsibilityLocalIT"
        End If
        End If
        End If


        'Zugriff auf Übersichts-Buttons regeln
        If ReadOnly = False Then
        ButtonShowOverView.Disabled = False
        ButtonEmailPreview.Disabled = False
        EditSubjectCase.Disabled = False
        Else
        ButtonShowOverView.Disabled = True
        ButtonEmailPreview.Disabled = True
        EditSubjectCase.Disabled = True
        End if
        'Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
        Dim Anfrageart
        Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0)
        If Anfrageart <> "RequestTypeContact" Then
        ComboIncidentStatus.Disabled = False
        Else
        ComboIncidentStatus.Disabled = True
        End If

        'Bei 2nd Level Dialog setzen der Benachrichtigung auf Email
        hlObj.SetValue "CaseGeneral.DefaultNotification",0,0,0,"DefaultNotificationEmail"
      
END SUB
SUB SearchAsset_AfterExecute()
        'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check Asset search status to set the caption of the button
        If SearchAsset.GetSearchState = 3 Then
        SearchAsset.Caption = "Reset"
        Else
        SearchAsset.Caption = "Inventar"
        End If
      
END SUB
SUB SearchAsset_AfterReset()
        Set objO = SearchAsset.GetObject("product", False)
        Set objT = SearchAsset.GetObject("product", True)

        Call objT.SetValue("AssetGeneral.AssetName", 0, 0, 0, "")
        Call objT.SetValue("AssetGeneral.Hostname", 0, 0, 0, "")
        Call objT.SetValue("TrumpfAssetGeneral.CINumber", 0, 0, 0, "")

        'Prüft ob Anfrager Objekt nicht vorhanden ist
        'Check wether the Caller object exist
        If IsObject(hlCaller) = False Or hlCaller.objID = 0 Then
        Call hlObj.SetValue("CaseGeneral.CostCenter", 0, 0, 0, "")
        End If

        'Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check Asset search status to set the caption of the button
        If SearchAsset.GetSearchState = 3 Then
        SearchAsset.Caption = "Reset"
        Else
        SearchAsset.Caption = "Inventar"
        End If
      
END SUB
SUB SearchAsset_Click()
        Dim ReadOnly, NoProduct
        ReadOnly = True
        NoProduct = True

        'Wenn kein Inventar gefunden wurde, abbrechen
        'Cancel If no Asset was found
        If hlProduct.GetType() = "TEMPOBJECT" Then Exit Sub

        'Ermitteln der Locale ID für die Sprachauswahl
        'Selecting the Locale ID for the desired language
        lcid = hlSession.GetLocaleID
        LangID = hlSession.LangIDFromLCID(lcid)

        'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
        'First of all check whether the Case is write protected
        If hlObj.IsReadOnly("CaseGeneral.Subject",0)=0 Then
        ReadOnly=False
        End If

        'Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
        'Check wether the Caller object exist
        If IsObject(hlProduct) = True And EditHostname.Text <> "" Then
        NoProduct = False
        End If

        If ReadOnly = False Then
        'Setzen des Inventars
        'Setting the asset
        varString=""
        varAType = hlProduct.GetType ()
        If varAType = "DesktopComputer" Or varAType = "ServerComputer" Or varAType = "NotebookComputer" Or varAType = "Printer" Then
        If EditHostname.Text <> "" Then varString = EditHostname.Text
        If EditAssetModel.Text<>"" Then varString = varString & " " & EditAssetModel.Text
        Else
        If EditAssetModel.Text<>"" Then varString = EditAssetModel.Text Else EditAssetModel.Text=" "
        End If
        EditAssetModel.Text = varString
        End If
      
END SUB
SUB SearchCaller_AfterExecute()
        'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check requester search status to set the caption of the button
        If SearchCaller.GetSearchState = 3 Then
        SearchCaller.Caption = "Reset"
        Else
        SearchCaller.Caption = "Search"
        End If

        'VIP-Status des Anfragers abfragen und Imp Vorgang setzen
        Valid = hlCaller.HasContent("PersonGeneral.VIPLevel",0,0)
        If Valid = 1 Then
        VIP = hlCaller.GetValue("PersonGeneral.VIPLevel",0,0,0,0)
        'If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
        Select Case vip
        Case "VIPLevelVIP"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,1
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(142,139,254)
        Case "VIPLevelITAdminDitzingen"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,2
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelITAdminTG"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,3
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelSAPKeyUserTUS"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,4
        ComboVIPStatus.Disabled = True
        Person.BackColor = RGB(205,250,255)
        Case "VIPLevelNon"
        ComboVIPStatus.Disabled = False
        ComboVIPStatus.SelectItem 0,0
        ComboVIPStatus.Disabled = True
        Person.BackColor = ""
        End Select
        End If

        sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0)
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        Dim tempmail
        tempmail = EditEmailAddress.text
        'Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
        'Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
        Dim strIncStatus:strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        strEmail = ""
        CallerCount = 0
        CallerCount = hlObj.GetItemCount(&H00000,130)

        If CallerCount > 0 Then
        Dim CaseCallers : Set CaseCallers = Nothing
        CaseCallers = hlObj.GetItems(&H00000,-1,-1,130)
        For Each Caller In CaseCallers
        CallerType = Caller.GetType
        If CallerType = "Employee" Then
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        If mailadr <> "" Then
        strEmail = strEmail + mailadr + ";"
        End If
        End If
        Next

        Else
        strEmail = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End if

        If InStr(strEmail,tempmail) > 0 Then
        Else
        strEmail = tempmail + ";" + strEmail
        End If

        If strEmail = "" Then
        strEmail = hlObj.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End If
        If strEmail = "-" Then
        strEmail = ""
        End If
        If sendmail = "EmailCallerYes" Then
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        TextBoxEmailTo.Required = True
        TextBoxEmailSubject.Required = True
        GroupBoxEmail.Disabled = False
        Else
        hlObj.SetValue "EmailSUAttribute.EmailSearchName",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailSearchResult",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailCC",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT",0,0,0,""
        GroupBoxEmail.Disabled = True
        TextBoxEmailTo.Required = False
        TextBoxEmailSubject.Required = False
        End if
      
END SUB
SUB SearchCaller_AfterReset()
        Set objO = SearchCaller.GetObject("caller", False)
        Set objT = SearchCaller.GetObject("caller", True)

        Call objT.SetValue("PersonGeneral.PersonSurname", 0, 0, 0, "")
        Call objT.SetValue("PersonGeneral.PersonGivenName", 0, 0, 0, "")
        Call objT.SetValue("PersonInformation.PersonOrganisation", 0, 0, 0, "")
        Call objT.SetValue("PersonInformation.PhoneNumber", 0, 0, 0, "")
        Call hlObj.SetValue("CaseGeneral.CostCenter", 0, 0, 0, "")


        'Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
        'Check requester search status to set the caption of the button
        If SearchCaller.GetSearchState = 3 Then
        SearchCaller.Caption = "Reset"
        Else
        EditSurname.Text = ""
        SearchCaller.Caption = "Search"
        End If

        'VIP-Status zurücksetzen
        ComboVIPStatus.SelectItem 0,0
        Person.BackColor = RGB(248,245,240)
      
END SUB
SUB SearchCaller_Click()
        Dim ReadOnly
        ReadOnly=True

        'Wenn keine Person gefunden wurde, abbrechen
        'Cancel If no person was found
        If hlCaller.GetType() = "TEMPOBJECT" Then Exit Sub

        'Zunächst überprüfen ob der Vorgang schreibgeschützt ist
        'First of all check whether the Case is write protected
        If hlObj.IsReadOnly("CaseGeneral.Subject",0)=0 Then
        ReadOnly=False
        End If

        'Ermitteln der Locale ID für die Sprachauswahl
        'Selecting the Locale ID for the desired language
        lcid = hlSession.GetLocaleID
        LangID = hlSession.LangIDFromLCID(lcid)




      
END SUB
SUB SetProblemText2Subject()
        varSubject = Left (EditProblem.Text, 100)
        If EditSubjectCase.Text="" Then EditSubjectCase.Text = replace(varSubject,Chr(13)&Chr(10)," ")
      
END SUB
SUB ComboIncidentStatus_SelectionChanged()
        Dim tempmail
        tempmail = EditEmailAddress.text
        Dim strIncStatus:strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        strEmail = ""
        CallerCount = 0
        CallerCount = hlObj.GetItemCount(&H00000,130)

        If CallerCount > 0 Then
        Dim CaseCallers : Set CaseCallers = Nothing
        CaseCallers = hlObj.GetItems(&H00000,-1,-1,130)
        For Each Caller In CaseCallers
        CallerType = Caller.GetType
        If CallerType = "Employee" Then
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        If mailadr <> "" Then
        strEmail = strEmail + mailadr + ";"
        End If
        End If
        Next
        Else
        strEmail = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End if

        If InStr(strEmail,tempmail) > 0 Then
        Else
        strEmail = tempmail + ";" + strEmail
        End If

        If strEmail = "" Then
        strEmail = hlObj.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End If
        If strEmail = "-" Then
        strEmail = ""
        End If
        Select Case strIncStatus
        Case "IncidentStatusSolved"
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerYes"
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        hlObj.SetValue "SUINFO.PUBLISHED",0,0,0,"1"
        GroupBoxEmail.Disabled = False
        LabelEmailBody.TextColor = "Red"
        ComplexTextEmailBody.Required = True
        TextBoxEmailTo.Required = True
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = True
        Case "IncidentStatusClosed"
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        hlObj.SetValue "SUINFO.PUBLISHED",0,0,0,"1"
        GroupBoxEmail.Disabled = False
        LabelEmailBody.TextColor = "Red"
        ComplexTextEmailBody.Required = True
        If strEmail = "" Then
        TextBoxEmailTo.Required = False
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerNo"
        Else
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerYes"
        TextBoxEmailTo.Required = True
        End If
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = True
        Case "IncidentStatusTimephased"
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerYes"
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        hlObj.SetValue "SUINFO.PUBLISHED",0,0,0,"1"
        GroupBoxEmail.Disabled = False
        LabelEmailBody.TextColor = "Red"
        ComplexTextEmailBody.Required = True
        TextBoxEmailTo.Required = True
        EditResubmissionTime.Required = True
        EditResubmissionTime.Disabled = false
        ComboBoxEmailCaller.Disabled = True
        Case "IncidentStatusWaitingforCustomer"
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerYes"
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        hlObj.SetValue "SUINFO.PUBLISHED",0,0,0,"1"
        GroupBoxEmail.Disabled = False
        LabelEmailBody.TextColor = "Red"
        ComplexTextEmailBody.Required = True
        TextBoxEmailTo.Required = True
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = True
        Case "IncidentStatusWaitingforExtern"
        LabelEmailBody.TextColor = "Black"
        ComplexTextEmailBody.Required = False
        TextBoxEmailTo.Required = False
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = False
        Case "IncidentStatusToProof"
        LabelEmailBody.TextColor = "Black"
        ComplexTextEmailBody.Required = False
        TextBoxEmailTo.Required = False
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = False
        Case "IncidentStatusRouted"
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerNo"
        LabelEmailBody.TextColor = "Black"
        ComplexTextEmailBody.Required = False
        TextBoxEmailTo.Required = False
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = False
        Case "IncidentStatusNew"
        LabelEmailBody.TextColor = "Black"
        ComplexTextEmailBody.Required = False
        TextBoxEmailTo.Required = False
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = False
        Case "IncidentStatusInProgress"
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,"EmailCallerNo"
        LabelEmailBody.TextColor = "Black"
        ComplexTextEmailBody.Required = False
        TextBoxEmailTo.Required = False
        EditResubmissionTime.Required = False
        EditResubmissionTime.Disabled = true
        EditResubmissionTime.DeleteContent
        ComboBoxEmailCaller.Disabled = False
        End Select

      
END SUB
SUB ComboRequestType_SelectionChanged()
        Dim Anfrageart, Status
        Anfrageart = hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0)
        Status = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)

        ComboProductionalRelevanz.Disabled = false

        If Anfrageart <> "RequestTypeIncident" Then
        ComboImpact.Disabled = True
        ComboFunctionalRange.Disabled = True
        hlObj.SetValue "CaseClassificationAttribute.Impact",0,0,0,"ImpactOne"
        hlObj.SetValue "IncidentAttribute.FunctionalRange",0,0,0,"FunctionalRangePartFailure"
        hlObj.SetValue "IncidentAttribute.ProductionalRelevanz",0,0,0,"ProductionalRelevanzAdministrativeProcess"
        Else
        ComboImpact.Disabled = False
        ComboFunctionalRange.Disabled = False
        hlObj.SetValue "IncidentAttribute.ProductionalRelevanz",0,0,0,"ProductionalRelevanzSupportProcess"
        End If

        If Anfrageart <> "RequestTypeContact" Then
        CaseProblem.Disabled = False
        If Status <> "IncidentStatusClosed" Then
        ComboBoxEmailCaller.Disabled = False
        Else
        ComboBoxEmailCaller.Disabled = True
        End If
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        Else
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
        End If


      
END SUB
SUB OnSave()
        'Priorität leeren, damit globale SLA´s auch runterstufen können
        hlObj.SetValue "CaseClassificationAttribute.Priority",0,0,0,"Priority5"

        CheckOverView = ""
        CheckOverView = hlObj.GetValue("CaseGeneral.Overview",0,0,0,0)
        If CheckOverView <> "" Then
        hlObj.SetValue "CaseGeneral.Overview",0,0,0,""
        End If
        CheckSummaryHTML = ""
        CheckSummaryHTML = hlObj.GetValue("CaseGeneral.SummaryHTML.TEXTVALUE",0,0,0,0)
        If CheckSummaryHTML <> "" Then
        hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE",0,0,0,""
        hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT",0,0,0,""
        'Button "Übersicht" entsperren
        ButtonShowOverView.Disabled = False
        End if





      
END SUB
SUB TreeKeyword_ondatachange()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Bitte zuerst das Ticket reservieren.")
        Else
        'Aktuellen Agent auslesen
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
        cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
        cn.ConnectionTimeout = 10
        cn.Open

        'Ditzingen oder TG auslesen
        Dim agentid,responsibility
        Set rs_resp = createobject("ADODB.Recordset")
        Set rs_resp = cn.Execute("Select responsibility from AgentID_responsibility where agentid = " & cstr(agent))
        responsibility = rs_resp.fields("responsibility").value
        rs_resp.close

        'Keyword einlesen
        Dim kw
        kw = hlObj.GetValue("Keywords.Keyword",0,0,0,1)
        If responsibility = 112545 Then
        'KeywordOrga Wert aus Vergleichstabelle einlesen
        Dim kwo
        Set rs_kwkwo = createobject("ADODB.Recordset")
        Set rs_kwkwo = cn.Execute("Select keywordorga from kw_kwo_mapping where keywordid = "& cstr(kw))
        Do While Not rs_kwkwo.EOF
        kwo = rs_kwkwo.fields("keywordorga").value
        rs_kwkwo.MoveNext
        Loop
        If Not kwo = "" Then
        hlObj.SetValue "Keywords.KeywordOrga",0,0,0,kwo
        TreeKeywordOrga.SelectTreeItem kwo
        End If
        rs_kwkwo.close
        Else
        'Wert für die TG setzen
        'Dim tg
        'tg = HIER TG Value einlesen
        'hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
        'TreeKeywordOrga.SelectTreeItem tg
        End If

        'Datenbankverbindung zu helpline_replication schließen
        cn.close
        Set cn = Nothing
        End If
      
END SUB
SUB ComboLevel_SelectionChanged()
        'Bei Änderung des Supportlevels automatisch den Status auf "Weitergeleitet" setzen
        Dim level
        level = hlObj.GetValue("IncidentAttribute.EscalationLevel",0,0,0,0)

        If level = "EscalationLevelLevel2" Then
        hlObj.SetValue "IncidentAttribute.IncidentStatus",0,0,0,"IncidentStatusRouted"
        End If
        If level = "EscalationLevelLevel1" Then
        hlObj.SetValue "IncidentAttribute.IncidentStatus",0,0,0,"IncidentStatusRouted"
        End if
      
END SUB
SUB ButtonDiscovery_Click()
        Dim Hostname : Hostname = hlProduct.getvalue("AssetGeneral.Hostname",0,0,0,0)
        Dim wshshell, oExec
        Set wshShell = CreateObject("Wscript.Shell")
        Command1 = "c:\program files\internet explorer\iexplore.exe http://srv01inv1/discovery/Reports/List.aspx?q=" + Hostname + "&flgDevice=1"
        Set oExec = wshShell.Exec(Command1)
      
END SUB
SUB b_template_save_Click()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Please reserve the ticket first.")

        Else

        'Templatenamen eingeben
        Dim name
        name = InputBox("Please type in a descriptive name for the template:","templatename","Maximum of 100 characters.")

        'Bei Abbruch nichts unternehmen, sonst weiter im Script
        IF name =FALSE THEN
        ELSE

        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
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
        result = MsgBox("Button YES => personal template for: " & agent_displayname & chr(10)&chr(13)&chr(13) & "or" & chr(10)&chr(13)&chr(13) & "Button NO => team template for: ''" & teamDisplayname & "''",4,"personal template or team template?")
        If result = 6 Then
        'Persönliches Insert auf Datenbank starten
        Set rs = cn.execute("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('" & cstr(agent) & "','" & name & "','" & hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0) &"','" & Replace(hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0), "'","''") &"','" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText",0,0,0,0), "'","''") &"','" & Replace(hlObj.GetValue("CaseSolution.SolutionText",0,0,0,0), "'","''") &"','" & hlObj.GetValue("Keywords.Keyword",0,0,0,0) &"','" & hlObj.GetValue("Keywords.KeywordOrga",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.EscalationLevel",0,0,0,0) &"','" & hlObj.GetValue("CaseClassificationAttribute.Impact",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.FunctionalRange",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0) &"','" & hlObj.GetValue("CaseGeneral.DefaultNotification",0,0,0,0) &"','" & cstr(agent) & "','" & hlObj.GetValue("IncidentAttribute.Convenience",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailTo",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailCC",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailSubject",0,0,0,0) &"')")
        Else
        'Team Insert auf Datenbank starten
        Set rs = cn.execute("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('" & cstr(teamID) & "','" & name & "','" & hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0) &"','" & Replace(hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0), "'","''") &"','" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText",0,0,0,0), "'","''") &"','" & Replace(hlObj.GetValue("CaseSolution.SolutionText",0,0,0,0), "'","''") &"','" & hlObj.GetValue("Keywords.Keyword",0,0,0,0) &"','" & hlObj.GetValue("Keywords.KeywordOrga",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.EscalationLevel",0,0,0,0) &"','" & hlObj.GetValue("CaseClassificationAttribute.Impact",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.FunctionalRange",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0) &"','" & hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0) &"','" & hlObj.GetValue("CaseGeneral.DefaultNotification",0,0,0,0) &"','" & cstr(agent) & "','" & hlObj.GetValue("IncidentAttribute.Convenience",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailTo",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailCC",0,0,0,0) &"','" & hlObj.GetValue("EmailSUAttribute.EmailSubject",0,0,0,0) &"')")

        End If
        'Verbindung schließen
        cn.close

        End if
        End if

      
END SUB
SUB b_template_load_Click()
        'Prüfen ob Template in der Checkbox ausgewählt wurde
        If cb_template_load.GetCurSel = -1 or l_templateID.text = "" then
        Dim msg
        msg = MsgBox("Please select a template from the list." & Chr(13) & Chr(10) & "If the list is empty, there is no template existing.",vbOKOnly,"No data record available.")
        Else

        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Angewählte ID aus Label auslesen
        Dim templateid
        templateid = l_templateID.Text

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
        cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
        cn.ConnectionTimeout = 10
        cn.Open

        'Inhalte von agent_templates in das Recordset einlesen
        Set rs = createobject("ADODB.Recordset")
        Set rs = cn.Execute("Select * from templater where template_id = " & templateid)

        hlObj.SetValue "IncidentAttribute.RequestType",0,0,0,rs.fields("Requesttype").value
        If hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0) = "" then
        hlObj.SetValue "CaseDescription.DescriptionText",0,0,0,rs.fields("descriptiontext").value
        Else
        End If
        hlObj.SetValue "CaseDiagnosis.DiagnosisText",0,0,0,rs.fields("diagnosistext").value
        hlObj.SetValue "CaseSolution.SolutionText",0,0,0,rs.fields("solutiontext").value
        hlObj.SetValue "Keywords.Keyword",0,0,0,rs.fields("keyword").value
        hlObj.SetValue "Keywords.KeywordOrga",0,0,0,rs.fields("keywordorga").value
        hlObj.SetValue "IncidentAttribute.EscalationLevel",0,0,0,rs.fields("EscalationLevel").value
        hlObj.SetValue "CaseClassificationAttribute.Impact",0,0,0,rs.fields("Impact").value
        hlObj.SetValue "IncidentAttribute.FunctionalRange",0,0,0,rs.fields("FunctionalRange").value
        hlObj.SetValue "IncidentAttribute.ProductionalRelevanz",0,0,0,rs.fields("ProductionalRelevance").value
        hlObj.SetValue "EmailSUAttribute.EmailCaller",0,0,0,rs.fields("EmailCaller").value
        hlObj.SetValue "IncidentAttribute.IncidentStatus",0,0,0,rs.fields("IncidentStatus").value
        hlObj.SetValue "CaseGeneral.DefaultNotification",0,0,0,rs.fields("DefaultNotification").value
        hlObj.SetValue "IncidentAttribute.Convenience",0,0,0,rs.fields("PCAssoziated").value
        hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,rs.fields("EmailBodytext").value
        hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT",0,0,0,rs.fields("EmailBodyRawtext").value
        'hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,rs.fields("EmailTo").value
        hlObj.SetValue "EmailSUAttribute.EmailCC",0,0,0,rs.fields("EmailCC").value
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        If hlObj.GetValue("EmailSUAttribute.EmailSubject",0,0,0,0) = "" then
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,rs.fields("EmailSubject").value
        End If

        'Subject Setzen
        varSubject = Left (EditProblem.Text, 100)
        If EditSubjectCase.Text="" Then EditSubjectCase.Text = replace(varSubject,Chr(13)&Chr(10)," ")

        'Übertrag der Caller in das An-Feld
        Dim tempmail
        tempmail = EditEmailAddress.text
        strEmail = ""
        CallerCount = 0
        CallerCount = hlObj.GetItemCount(&H00000,130)

        If CallerCount > 0 Then
        Dim CaseCallers : Set CaseCallers = Nothing
        CaseCallers = hlObj.GetItems(&H00000,-1,-1,130)
        For Each Caller In CaseCallers
        CallerType = Caller.GetType
        If CallerType = "Employee" Then
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        If mailadr <> "" Then
        strEmail = strEmail + mailadr + ";"
        End If
        End If
        Next
        Else
        strEmail = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End if

        If InStr(strEmail,tempmail) > 0 Then
        Else
        strEmail = tempmail + ";" + strEmail
        End If

        If strEmail = "" Then
        strEmail = hlObj.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End If
        If strEmail = "-" Then
        strEmail = ""
        End If

        'Aktivieren der Felder je nach EmailCaller Wert
        sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0)
        If sendmail = "EmailCallerYes" Then
        TextBoxEmailTo.Required = True
        TextBoxEmailSubject.Required = True
        GroupBoxEmail.Disabled = False
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        Else
        TextBoxEmailTo.Required = False
        TextBoxEmailSubject.Required = False
        GroupBoxEmail.Disabled = True
        End If

        'Aktivieren/Deaktivieren der Felder je nach gesetzter Anfrageart
        ComboProductionalRelevanz.Disabled = false
        If Anfrageart <> "RequestTypeIncident" Then
        ComboImpact.Disabled = True
        ComboFunctionalRange.Disabled = True
        hlObj.SetValue "CaseClassificationAttribute.Impact",0,0,0,"ImpactOne"
        hlObj.SetValue "IncidentAttribute.FunctionalRange",0,0,0,"FunctionalRangePartFailure"
        hlObj.SetValue "IncidentAttribute.ProductionalRelevanz",0,0,0,"ProductionalRelevanzAdministrativeProcess"
        Else
        ComboImpact.Disabled = False
        ComboFunctionalRange.Disabled = False
        hlObj.SetValue "IncidentAttribute.ProductionalRelevanz",0,0,0,"ProductionalRelevanzSupportProcess"
        End If

        If Anfrageart <> "RequestTypeContact" Then
        CaseProblem.Disabled = False
        Dim status
        status = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        If Status <> "IncidentStatusClosed" Then
        ComboBoxEmailCaller.Disabled = False
        Else
        ComboBoxEmailCaller.Disabled = True
        End If
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        Else
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
        End If

        'Recordset schließen
        rs.close
        Set rs = Nothing

        'Datenbankverbindung zu helpline_replication schließen
        cn.close
        Set cn = Nothing

        End If
      
END SUB
SUB b_template_change_Click()
        If cb_template_load.GetCurSel = -1 or l_templateID.text = "" then
        MsgBox("Please select template from list first.")
        Else

        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Angewählte ID aus Label auslesen
        Dim templateid
        templateid = l_templateID.Text

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
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
        If editor <> cstr(agent) then
        Dim msg2
        msg2 = MsgBox("You can only overwrite self-created templates." & chr(10)&chr(13) & "template: " & templateid & " was created by: " & agent_displayname & "",vbOKOnly,"Overwrite is not allowed")
        Else
        Dim name
        name = InputBox("Please type in a descriptive name: ","overwrite template",templatename)
        IF name=FALSE THEN
        ELSE

        'Abfrage ob Update erwünscht
        Dim result
        result = MsgBox("Möchten Sie das Template:  ''" & templatename & "''  überschreiben?",4,"Template überschreiben?")
        If result = 6 Then

        'Update auf Datenbank wird ausgeführt
        Set rs = cn.execute("Update templater set templatename = '" & name &"', Requesttype = '" & hlObj.GetValue("IncidentAttribute.RequestType",0,0,0,0) & "',descriptiontext = '" & Replace(hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0), "'","''") & "', diagnosistext = '" & Replace(hlObj.GetValue("CaseDiagnosis.DiagnosisText",0,0,0,0), "'","''") & "', solutiontext = '" & Replace(hlObj.GetValue("CaseSolution.SolutionText",0,0,0,0), "'","''") & "', keyword = '" & hlObj.GetValue("Keywords.Keyword",0,0,0,0) & "', keywordorga = '" & hlObj.GetValue("Keywords.KeywordOrga",0,0,0,0) &"', EscalationLevel = '" & hlObj.GetValue("IncidentAttribute.EscalationLevel",0,0,0,0) &"',Impact = '" & hlObj.GetValue("CaseClassificationAttribute.Impact",0,0,0,0) &"',FunctionalRange = '" & hlObj.GetValue("IncidentAttribute.FunctionalRange",0,0,0,0) &"',ProductionalRelevance = '" & hlObj.GetValue("IncidentAttribute.ProductionalRelevanz",0,0,0,0) &"',EmailCaller = '" & hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0) &"',IncidentStatus = '" & hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0) &"',DefaultNotification = '" & hlObj.GetValue("CaseGeneral.DefaultNotification",0,0,0,0) &"',editor = '" & cstr(agent) & "',PCAssoziated = '" & hlObj.GetValue("IncidentAttribute.Convenience",0,0,0,0) &"',EmailBodyRawtext = '" & hlObj.GetValue("EmailSUAttribute.EmailBody.Rawtext",0,0,0,0) &"',EmailBodytext = '" & hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0) &"',EmailTo = '" & hlObj.GetValue("EmailSUAttribute.EmailTo",0,0,0,0) &"',EmailCC = '" & hlObj.GetValue("EmailSUAttribute.EmailCC",0,0,0,0) &"',EmailSubject = '" & hlObj.GetValue("EmailSUAttribute.EmailSubject",0,0,0,0) &"' where template_id = " & cstr(templateid))
        Set rs = nothing
        Else
        End If

        'EndIF Überschreiben
        End If

        'EndIf Agent = Editor
        End if

        'Verbindung schließen
        cn.close

        'EndIf Wurde ein Checkbox-Wert zuvor angewählt
        END IF
      
END SUB
SUB cb_template_load_onfocus()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Please reserve the ticket first.")
        EditSurname.RequestFocus = true
        Else

        'Vorhandene Checkbox Werte entfernen
        cb_template_load.ResetContent

        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
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
        Set rs = cn.Execute("Select template_id,templatename from templater where agentid = " & cstr(agent) & " order by agentid, cast(Templatename as varchar(500))" )
        On Error Resume Next
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
        Set rs2 = cn.Execute("Select template_id,templatename from templater where agentid = " & cstr(teamID) & " order by agentid, cast(Templatename as varchar(500))" )
        On Error Resume Next
        rs2.MoveFirst
        Do While Not rs2.eof
        cb_template_load.AddItem(rs2.fields("templatename").value)
        anzahl_team_templates = anzahl_team_templates +1
        rs2.MoveNext
        Loop

        'Recordset schließen
        rs.close
        rs2.close


        'Datenbankverbindung zu helpline_replication schließen
        cn.close
        Set cn = Nothing

        End If

      
END SUB
SUB b_template_delete_Click()
        'Prüfen ob Template in der Checkbox ausgewählt wurde
        If cb_template_load.GetCurSel = -1 or l_templateID.text = "" then
        Dim msg
        msg = MsgBox("Please select a template from the list." & Chr(13) & Chr(10) & "If the list is empty, there is no template existing.",vbOKOnly,"No data record available.")

        Else

        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Angewählte ID aus Label auslesen
        Dim templateid
        templateid = l_templateID.Text

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
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

        If editor <> cstr(agent) Then
        Dim msg2
        msg2 = MsgBox("You are only allowed to delete self-created tickets." & chr(10)&chr(13) & "Template ID: " & templateid & " was created by:" & agent_displayname & "",vbOKOnly,"Delete not allowed.")
        Else

        'Abfrage ob Löschen erwünscht
        Dim result
        result = MsgBox("Do you really want to delete the template?",4,"Delete template?")
        If result = 6 Then

        'Zeile von agent_templates löschen
        Set rs = createobject("ADODB.Recordset")
        Set rs = cn.Execute("Delete from templater where template_id = " & cstr(templateid))

        'Auswahl der Checkbox zurücksetzen und ID auf Null
        cb_template_load.ResetContent
        l_templateid.text = ""

        'Recordset schließen
        Set rs = Nothing
        Else
        End If


        'End If Editor = Agent
        End If

        'Datenbankverbindung zu helpline_replication schließen
        cn.close
        Set cn = Nothing

        'Vorhandene Checkbox Werte entfernen
        cb_template_load.ResetContent
        l_templateID.Text = ""

        End If

      
END SUB
SUB cb_template_load_SelectionEndOK()
        'Agentid auslesen anhand des aktuellen Agenten
        Dim agent, team
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Angewählte Position bestimmen
        Dim position
        position = cb_template_load.GetCurSel +1

        'Datenbankverbindung zu helpline_replication
        Set cn=CreateObject("ADODB.Connection")
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
        On Error Resume Next
        rs_anzahl.MoveFirst
        Do While Not rs_anzahl.eof
        anzahl_agent_templates = anzahl_agent_templates + 1
        rs_anzahl.MoveNext
        Loop
        rs_anzahl.close

        If position =< anzahl_agent_templates Then
        'Select für Agententemplate ausführen
        Set rs_agent = createobject("ADODB.Recordset")
        Set rs_agent = cn.Execute("Select template_id from templater where agentid = '" & cstr(agent) & "' order by agentid, cast(Templatename as varchar(500))" )
        On Error Resume Next
        rs_agent.MoveFirst
        For i = 1 To position
        l_templateID.Text = rs_agent.fields("template_id").value
        rs_agent.MoveNext
        Next
        'Dataset schließen
        rs_agent.close

        Else

        'Prüfung, ob Trennlinie ausgewählt wurde.
        If cb_template_load.GetCurSel = anzahl_agent_templates then
        l_templateID.Text = ""
        'cb_template_load.ResetContent

        Else
        'Select für Teamtemplate ausführen  - "Position -1" wegen Trennzeile zwischen Templatetypen
        position = position - anzahl_agent_templates - 1
        Set rs_team = createobject("ADODB.Recordset")
        Set rs_team = cn.Execute("Select template_id from templater where agentid = '" & cstr(teamID) & "' order by agentid, cast(Templatename as varchar(500))" )
        On Error Resume Next
        rs_team.MoveFirst
        For i = 1 To position
        l_templateID.Text = rs_team.fields("template_id").value
        rs_team.MoveNext
        Next
        'Dataset schließen
        rs_team.close

        End If
        End If

        'DB schließen
        cn.close
      
END SUB
SUB ButtonSCCMRemote_Click()
        Dim wshshell, oExec,OsType
        Set wshShell = CreateObject("Wscript.Shell")

        'Ermitteln der Locale ID für die Sprachauswahl
        'Selecting the Locale ID for the desired language
        lcid = hlSession.GetLocaleID
        LangID = hlSession.LangIDFromLCID(lcid)

        If hlObj.IsReadOnly("CASEINFO.REACTIONTIME",0)=0 Then

        objType = hlProduct.GetType
        If objType = "DesktopComputer" Or objType = "ServerComputer" Or objType = "NotebookComputer" Then
        'Auslesen des gewählten Computers
        'Reading the selected computer
        host = EditHostname.Text

        If host <> "" Then
        On Error Resume Next
        'Kommandozeile für den Aufruf von On Command Remote Master
        'Command lin for calling On Command Remote Master
        'Command1="""%programfiles%"\smsadmin\bin\i386\remote.exe 2 "" & host
        OsType = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
        If OsType = 32 then
        'x86
        Command1="c:\Program Files\Microsoft Configuration Manager Console\AdminUI\bin\i386\rc.exe 1 " & host
        else
        'x64
        Command1="c:\Program Files (x86)\Microsoft Configuration Manager Console\AdminUI\bin\i386\rc.exe 1 " & host
        end if

        RemoteTool = "SCCM Remote"

        Set oExec = wshShell.Exec(Command1)
        If err.Number = -2147024893 Then
        If LangID = 7 Then
        msgbox "Auf Ihrem Computer ist das Remote Tool " & RemoteTool & " nicht installiert." & vbLf & "Bitte wenden Sie sich an Ihren Administrator.",vbExclamation,"helpLine - ClassicDesk"
        Else
        msgbox "The remote tool " & RemoteTool & " is not installed on your computer." & vbLf & "Please consult your administrator.",vbExclamation,"helpLine - ClassicDesk"
        End If
        End If
        End If
        Else
        If LangID = 7 Then
        msgbox "Es wurde kein Computer als Inventar ausgewählt." & vbLf & "Bitte wählen Sie einen Computer für den Vorgang aus.", vbExclamation, "helpLine - ClassicDesk"
        Else
        msgbox "No computer has been selected." & vbLf & "Please select a computer for this Case.", vbExclamation, "helpLine - ClassicDesk"
        End If
        End If
        End If



      
END SUB
SUB ButtonShowOverView_Click()
        'Ermitteln der Locale ID für die Sprachauswahl
        'Selecting the Locale ID for the desired language
        lcid = hlSession.GetLocaleID
        LangID = hlSession.LangIDFromLCID(lcid)

        CaseOwner = hlObj.GetValue("HLOBJECTINFO.OWNER",0,0,0,0)
        Agent = ""
        If LangID = 7 Then
        Problemtitle = "<b>====== Problembeschreibung ======" & " [von Agent : " & CaseOwner & "]</b>" & vbNewLine
        Diagnosistitle = "<b>====== Kommunikation ======</b>" & vbNewLine
        Solutiontitle = "<b>====== Lösungsbeschreibung ======" & " [von Agent : " & hlObj.GetValue("SUINFO.EDITOR", 0, 0, 0, 0) & "]</b>" & vbNewLine
        Else
        Problemtitle = "<b>====== Problemdescription ======" & " [by Agent : " & CaseOwner & "]</b>" & vbNewLine
        Diagnosistitle = "<b>====== Diagnosisactivities ======</b>" & vbNewLine
        Solutiontitle = "<b>====== Final solution ======" & " [by Agent : " & hlObj.GetValue("SUINFO.EDITOR", 0, 0, 0, 0) & "]</b>" & vbNewLine
        End If
        'VG-Beschreibung
        DescrText = ""
        DescrText = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,1,0)
        If DescrText = "" Then
        DescrText = hlObj.GetValue("CaseDescription.DescriptionText",0,0,0,0)
        End if
        If DescrText <> "" Then
        DescrText = Replace(DescrText, vbCrLf, "<br>")
        ProblemAll = Problemtitle & DescrText & vbNewLine
        End If
        'VG-Lösung
        'nur bei Status "Geschlossen" aus der aktuellen SU den Text holen
        actStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        SolText = ""
        If actStatus = "IncidentStatusClosed" Then
        SolText = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0)
        End If
        If SolText = "" Then
        SolText = hlObj.GetValue("CaseSolution.SolutionText",0,0,0,0)
        End If
        If SolText <> "" Then
        SolText = Replace(SolText, vbCrLf, "<br>")
        SolutionAll = Solutiontitle & SolText
        End if

        SUIdx = hlObj.GetValue("SUINFO.INDEX",0,0,0,0)
        If SUIdx > 0 Then
        'Pro SU prüfen, ob Tätigkeitsbeschreibung eingetragen ist
        For i=1 To SUIdx
        SUDiagnosisIntern = "<b> --- intern --- </b>"
        SUDiagnosis = ""
        SUDiagnosis = hlObj.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, i, 0)
        'SUDiagnosis = Replace(SUDiagnosis, Chr(13) & Chr(10), " ")
        SUDiagnosis = Replace(SUDiagnosis, vbCrLf, "<br>")
        SUDiagnosisExtern = "<b> --- extern --- </b>"
        SUDiagnosisExt = ""
        SUDiagnosisExt = hlObj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, i, 0)
        If SUDiagnosis <> "" Then
        SUActivity = hlObj.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, i, 0)
        SURegTime = hlObj.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, i, 0)
        Agent = hlObj.GetValue("SUINFO.EDITOR", 0, 0, i, 0)
        DiagnosisAll = DiagnosisAll & SUDiagnosisIntern & vbNewLine & "<b>" & i & ". SU (" & Agent & ") -> " & SUActivity & " [" & SURegTime & "]:" & "</b>" & vbNewLine & SUDiagnosis & vbNewLine & String(80, "-") & vbNewLine
        End If
        If SUDiagnosisExt <> "" Then
        'SUDiagnosisExt = Replace(SUDiagnosisExt, vbCrLf, "<br>")
        SUActivity = hlObj.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, i, 0)
        SURegTime = hlObj.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, i, 0)
        Agent = hlObj.GetValue("SUINFO.EDITOR", 0, 0, i, 0)
        DiagnosisAll = DiagnosisAll & SUDiagnosisExtern & vbNewLine & "<b>" & i & ". SU (" & Agent & ") -> " & SUActivity & " [" & SURegTime & "]:" & "</b>" & vbNewLine & SUDiagnosisExt & vbNewLine & String(80, "-") & vbNewLine
        End If
        Next
        End If
        If DiagnosisAll <> "" Then
        DiagnosisAll = Diagnosistitle & DiagnosisAll
        End If
        ProblemAll = ProblemAll & DiagnosisAll & SolutionAll
        'hlObj.SetValue "CaseGeneral.Overview",0,0,0,ProblemAll
        ProblemAll = Replace(ProblemAll, vbCrLf, "<br>")
        hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE",0,0,0,ProblemAll

        'Button nach 1. Klick sperren
        'ButtonShowOverView.Disabled = True
      
END SUB
SUB ComboBoxEmailCaller_SelectionChanged()
        sendmail = hlObj.GetValue("EmailSUAttribute.EmailCaller",0,0,0,0)
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        Dim tempmail
        tempmail = EditEmailAddress.text
        'Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
        'Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
        Dim strIncStatus:strIncStatus = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        strSubject = hlObj.GetValue("CaseGeneral.Subject",0,0,0,0)
        strEmail = ""
        CallerCount = 0
        CallerCount = hlObj.GetItemCount(&H00000,130)

        If CallerCount > 0 Then
        Dim CaseCallers : Set CaseCallers = Nothing
        CaseCallers = hlObj.GetItems(&H00000,-1,-1,130)
        For Each Caller In CaseCallers
        CallerType = Caller.GetType
        If CallerType = "Employee" Then
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        If mailadr <> "" Then
        strEmail = strEmail + mailadr + ";"
        End If
        End If
        Next

        Else
        strEmail = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End if

        If InStr(strEmail,tempmail) > 0 Then
        Else
        strEmail = tempmail + ";" + strEmail
        End If

        If strEmail = "" Then
        strEmail = hlObj.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End If
        If strEmail = "-" Then
        strEmail = ""
        End If
        If sendmail = "EmailCallerYes" Then
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,strSubject
        TextBoxEmailTo.Required = True
        TextBoxEmailSubject.Required = True
        GroupBoxEmail.Disabled = False
        Else
        hlObj.SetValue "EmailSUAttribute.EmailSearchName",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailSearchResult",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailCC",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailSubject",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,""
        hlObj.SetValue "EmailSUAttribute.EmailBody.RAWTEXT",0,0,0,""
        GroupBoxEmail.Disabled = True
        TextBoxEmailTo.Required = False
        TextBoxEmailSubject.Required = False
        End if
      
END SUB
SUB ButtonSearchMail_Click()
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

        If Name <> "" Then
        '------------------------------------------------------------------------------------------------
        'Ermitteln der Email-Adressen auf Bases des eingegebenen Namens
        Set cn2=createobject("ADODB.Connection")

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
        i =  i + 1
        ComboBoxEmailSearchResult.AddItem rs2.fields(0).value
        If i = 1 Then
        ComboBoxEmailSearchResult.Text = rs2.fields(0).value
        End if
        rs2.movenext
        Loop
        'Verbindung schließen
        rs2.close
        cn2.close

        End If





      
END SUB
SUB ButtonTo_Click()
        email = ComboBoxEmailSearchResult.Text
        Recipient = TextBoxEmailTo.Text
        If email = "" Then
        MsgBox "Bitte eine Email-Adresse auswählen!"
        Else
        fullemailstring = len(email)
        pos=Instr(1,email,":",1)
        emailstring=clng(fullemailstring)-clng(pos)
        email=Right(email,CLNG(emailstring))
        If Recipient = "" Then
        Recipient = email
        Else
        If RIGHT(Recipient,1) = ";" Then
        Recipient = Recipient+email
        Else
        Recipient = Recipient+";"+email
        End If
        End If
        TextBoxEmailTo.Text = Recipient
        End If
      
END SUB
SUB ButtonCC_Click()
        email = ComboBoxEmailSearchResult.Text
        RecipientCC = TextBoxEmailCC.Text
        If email = "" Then
        MsgBox "Bitte eine Email-Adresse auswählen!"
        Else
        fullemailstring = len(email)
        pos=Instr(1,email,":",1)
        emailstring=clng(fullemailstring)-clng(pos)
        email=Right(email,CLNG(emailstring))
        If RecipientCC = "" Then
        RecipientCC = email
        Else
        If RIGHT(RecipientCC,1) = ";" Then
        RecipientCC = RecipientCC+email
        Else
        RecipientCC = RecipientCC+";"+email
        End If
        End If
        TextBoxEmailCC.Text = RecipientCC
        End If

      
END SUB
SUB ButtonSetAgent1_Click()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Bitte zuerst das Ticket reservieren.")
        Else

        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_data
        Set cn=CreateObject("ADODB.Connection")
        cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
        cn.ConnectionTimeout = 10
        cn.Open

        'Teamname auslesen
        Dim agentid,internalname
        Set rs_kwo = createobject("ADODB.Recordset")
        Set rs_kwo = cn.Execute("Select name,internalname from vw_agent_to_first_keywordorga where agentid = " & cstr(agent))
        internalname = rs_kwo.fields("internalname").value

        'Wert in Schlagwort schreiben
        hlObj.SetValue "Keywords.KeywordOrga",0,0,0,internalname
        TreeKeywordOrga.SelectTreeItem internalname

        'Datenbankverbindung zu helpline_replication schließen
        rs_kwo.close
        cn.close
        Set cn = Nothing

        End If




      
END SUB
SUB ButtonSetKW_Click()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Bitte zuerst das Ticket reservieren.")
        Else
        'Aktuellen Agent auslesen
        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_replication
        Set cn1=CreateObject("ADODB.Connection")
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
        hlObj.SetValue "Keywords.Keyword",0,0,0,keywordid
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
        kw = hlObj.GetValue("Keywords.Keyword",0,0,0,1)
        If responsibility = 112545 Then
        'KeywordOrga Wert aus Vergleichstabelle einlesen
        Dim kwo
        Set rs_kwkwo = createobject("ADODB.Recordset")
        Set rs_kwkwo = cn1.Execute("Select keywordorga from kw_kwo_mapping where keywordid = "& cstr(kw))
        Do While Not rs_kwkwo.EOF
        kwo = rs_kwkwo.fields("keywordorga").value
        rs_kwkwo.MoveNext
        Loop
        If Not kwo = "" Then
        hlObj.SetValue "Keywords.KeywordOrga",0,0,0,kwo
        TreeKeywordOrga.SelectTreeItem kwo
        End If
        rs_kwkwo.close
        Else
        'Wert für die TG setzen
        'Dim tg
        'tg = HIER TG Value einlesen
        'hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
        'TreeKeywordOrga.SelectTreeItem tg
        End If

        'Datenbankverbindung zu helpline_replication schließen
        cn1.close
        Set cn1 = Nothing
        End If




      
END SUB
SUB ButtonResetTo_Click()
        CallerCount = 0
        CallerCount = hlObj.GetItemCount(&H00000,130)
        If CallerCount > 0 Then
        Dim CaseCallers : Set CaseCallers = Nothing
        CaseCallers = hlObj.GetItems(&H00000,-1,-1,130)
        For Each Caller In CaseCallers
        CallerType = Caller.GetType
        If CallerType = "Employee" Then
        mailadr = ""
        mailadr = Caller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        If mailadr <> "" Then
        strEmail = strEmail + mailadr + ";"
        End If
        End If
        Next
        Else
        strEmail = hlCaller.GetValue("PersonInformation.EmailAddress",0,0,0,0)
        End if

        Dim tempmail
        tempmail = EditEmailAddress.text
        If InStr(strEmail,tempmail) > 0 Then
        Else
        strEmail = tempmail + ";" + strEmail
        End If

        hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,strEmail
      
END SUB
SUB ButtonEmailPreview_Click()
        status = hlObj.GetValue("IncidentAttribute.IncidentStatus",0,0,0,0)
        HLinkToCase = "http://srv01itsm2/helpLinePortal"
        HTicketID = hlobj.GetValue("CASEINFO.REFERENCENUMBER",0,0,0,0)
        SubjectCase = hlobj.GetValue("EmailSUAttribute.EmailSubject",0,0,0,0)
        LanguageDE = 0
        MailTo = hlobj.GetValue("EmailSUAttribute.EmailTo",0,0,0,0)
        For z = 1 To len(MailTo)
        IF Mid(MailTo, z, 1) = "@" THEN
        CounterEmpf = CounterEmpf + 1
        End If
        Next
        If IsObject(hlCaller) = True Then
        surname = hlCaller.GetValue("PersonGeneral.PersonSurname",0,0,0,0)
        letteraddress = hlCaller.GetValue("PersonGeneral.ShortLetterAddress",0,0,0,0)
        language = hlCaller.GetValue("PersonGeneral.Language",0,0,0,0)
        If language <> "LanguageGerman" Then
        LanguageDE = -1
        Else
        LanguageDE = 1
        End If
        Else
        surname = "Unbekannt/Unknown"
        End If
        Editor = hlobj.GetValue("SUINFO.EDITOR",0,0,0,0)
        '----------------------------------------------------------------------------------------------------------
        'M.Rettig, 14.05.2012 - SU-Email als HTML-Vorschau
        If status = "IncidentStatusClosed" Then
        Const ForReading = 1, ForWriting = 2, ForAppending = 8
        Dim OriginDescr : OriginDescr = hlobj.GetValue("CaseDescription.DescriptionText",0,0,0,0)
        MailBody = hlobj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0)
        'Deutsche Werte
        If LanguageDE > 0 Then
        If letteraddress = "" Then
        letteraddress = "Herr/Frau"
        End If

        'Konstante Werte deutsch setzen
        TTicketID = "Ticketnummer"
        TStatus = "Status"
        HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus",7,0,LastSUIdx,0)
        TEditor = "Bearbeiter"
        TSubject = "Betreff:"
        If CounterEmpf > 1 Then
        Anrede = "Sehr geehrte "
        surname = "Damen und Herren"
        Else
        Anrede = "Sehr geehrte(r) " & CStr(letteraddress)
        End If
        TSolution = "Lösung:"
        TBeschr = "Ticket-Beschreibung:"
        TComplimentary = "Mit freundlichen Grüßen,"
        TSignature = "Ihr Team IT + Prozesse"
        TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!"
        Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME",7,0,0,0)
        Datum = Mid(Creationdate,1,10)
        subject = "Lösung zur IT Service Desk Anfrage " & " [#"
        subject = subject & HTicketID & "]" & " vom " & Datum
        TIntroduction = "Wir möchten Ihnen folgende Lösung übermitteln:"
        Else
        If letteraddress = "" Then
        letteraddress = "Mrs./Ms./Mr."
        End If

        'Konstante Werte englisch setzen
        TTicketID = "Ticket number"
        TStatus = "Status"
        HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus",9,0,LastSUIdx,0)
        TEditor = "Editor"
        TSubject = "Subject:"
        If CounterEmpf > 1 Then
        Anrede = "Dear "
        surname = "Sir or Madam"
        Else
        Anrede = "Dear " & CStr(letteraddress)
        End If
        TSolution = "Solution:"
        TBeschr = "Ticket-Description:"
        TComplimentary = "Best regards,"
        TSignature = "Your support team IT + Processes"
        TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!"
        Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME",9,0,0,0)
        Datum = Mid(Creationdate,1,10)
        subject = "Your support request from " & Datum & " with the reference no. [#"
        subject = subject & HTicketID & "]"
        TIntroduction = "We deliver to you the following solution description:"
        End If
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
        hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT",0,0,0,BodyText
        hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE",0,0,0,BodyText
        Else
        DiagnText = hlobj.GetValue("EmailSUAttribute.EmailBody.TEXTVALUE",0,0,0,0)
        If LanguageDE = 1 Then
        If letteraddress = "" Then
        letteraddress = "Herr/Frau"
        End If

        'Konstante Werte deutsch setzen
        TTicketID = "Ticketnummer"
        TStatus = "Status"
        HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus",7,0,LastSUIdx,0)
        TEditor = "Bearbeiter"
        TSubject = "Betreff:"
        If CounterEmpf > 1 Then
        Anrede = "Sehr geehrte "
        surname = "Damen und Herren"
        Else
        Anrede = "Sehr geehrte(r) " & CStr(letteraddress)
        End If
        TDiagnosis = "Zwischenbescheid"
        TResubTime = "Wiedervorlagedatum:"
        TComplimentary = "Mit freundlichen Grüßen,"
        TSignature = "Ihr Team IT + Prozesse"
        TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!"

        'Hier wird die Betreffzeile erstellt
        'The subject field is entered here
        Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME",7,0,0,0)
        Datum = Mid(Creationdate,1,10)
        ResubmissionTime = hlobj.GetValue("CASEINFO.RESUBMISSIONTIME",7,0,0,0)
        If ResubmissionTime <> "" Then
        If DateDiff("d",Now,ResubmissionTime) > 0 Then
        'If ResubmissionTime > Now Then
        ResubDatum = MID(ResubmissionTime,1,10)
        Else
        ResubDatum = ""
        End If
        End If
        subject = "Zwischenbescheid zur IT Service Desk Anfrage " & " [#"
        subject = subject & HTicketID & "]" & " vom " & Datum
        TIntroduction = "Wir möchten Ihnen folgende Nachricht übermitteln:"
        Else
        If letteraddress = "" Then
        letteraddress = "Mrs./Ms./Mr."
        End If

        'Konstante Werte englisch setzen
        TTicketID = "Ticket number"
        TStatus = "Status"
        HStatus = hlobj.GetValue("IncidentAttribute.IncidentStatus",9,0,LastSUIdx,0)
        TEditor = "Editor"
        TSubject = "Subject:"
        If CounterEmpf > 1 Then
        Anrede = "Dear "
        surname = "Sir or Madam"
        Else
        Anrede = "Dear " & CStr(letteraddress)
        End If

        TDiagnosis = "Intermediate Reply"
        TResubTime = "Resubmissiontime:"
        TComplimentary = "Best regards,"
        TSignature = "Your support team IT + Processes"
        TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!"


        'Hier wird die Betreffzeile erstellt
        'The subject field is entered here
        Creationdate = hlobj.GetValue("HLOBJECTINFO.CREATIONTIME",9,0,0,0)
        Datum = Mid(Creationdate,1,10)
        ResubmissionTime = hlobj.GetValue("CASEINFO.RESUBMISSIONTIME",9,0,0,0)
        If ResubmissionTime <> "" Then
        If DateDiff("d",Now,ResubmissionTime) > 0 Then
        'If ResubmissionTime > Now Then
        ResubDatum = MID(ResubmissionTime,1,10)
        Else
        ResubDatum = ""
        End If
        End If
        subject = "Your support request from " & Datum & " with the reference no. [#"
        subject = subject & HTicketID & "]"
        TIntroduction = "We deliver to you the following processing description:"
        End If

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
        If ResubDatum <> "" Then
        BodyText = replace(BodyText, "[$ResubmissionTime_Titel$]", TResubTime)
        BodyText = replace(BodyText, "[$ResubmissionTime$]", ResubDatum)
        Else
        BodyText = replace(BodyText, "[$ResubmissionTime_Titel$]", "")
        BodyText = replace(BodyText, "[$ResubmissionTime$]", "")
        End if
        BodyText = replace(BodyText, "[$ComplimentaryClose$]", TComplimentary)
        BodyText = replace(BodyText, "[$Signature$]", TSignature)
        'Schließt das File
        f.Close
        Set f = Nothing
        Set fso = Nothing
        hlObj.SetValue "CaseGeneral.SummaryHTML.RAWTEXT",0,0,0,BodyText
        hlObj.SetValue "CaseGeneral.SummaryHTML.TEXTVALUE",0,0,0,BodyText
        End If

      
END SUB
SUB ButtonSaveKW_Click()
        Dim isreserved
        isreserved = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,0)
        If isreserved = ""  Then
        MsgBox("Bitte zuerst das Ticket reservieren.")
        Else

        Dim agent
        agent = hlObj.GetValue("CASEINFO.RESERVEDBY",0,0,0,1)

        'Datenbankverbindung zu helpline_replication
        Set cn1=CreateObject("ADODB.Connection")
        cn1.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2"
        cn1.ConnectionTimeout = 10
        cn1.Open

        'Keyword einlesen und in Datenbank ablegen
        Dim personid,keywordid
        keywordid = hlObj.GetValue("Keywords.Keyword",0,0,0,1)
        If (CDbl(keywordid)) > 0 then
        'Personid über AgentID ermitteln
        Set rs_person = createobject("ADODB.Recordset")
        Set rs_person = cn1.Execute("Select personid from vw_Agent_Emplkeyword where agentid = " & cstr(agent))
        personid = rs_person.fields("personid").value
        rs_person.close

        'Datenbankverbindung zu helpline_data
        Set cn=CreateObject("ADODB.Connection")
        cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2"
        cn.ConnectionTimeout = 10
        cn.Open
        'Keyword schreiben
        Set rs_kw = createobject("ADODB.Recordset")
        Set rs_kw = cn.Execute("Update dbo.emplkeywords set keyword = " & cdbl(hlObj.GetValue("Keywords.Keyword",0,0,0,1)) & " where personid = " & cstr(personid))
        'Datenbank schließen
        'rs_kw.close
        cn.close
        Set cn = Nothing
        else
        MsgBox("Please select a keyword first.")
        End If


        'Datenbankverbindung zu helpline_replication schließen
        cn1.close
        Set cn1 = Nothing

        End If





      
END SUB
SUB EditSubjectCase_ondatachange()
        Dim Text
        If InStr (1, EditSubjectCase.Text, "Notfalltransport_SAP", vbTextCompare) then
        CaseProblem.Disabled = False
        CaseProblem.Disabled = False
        ComboBoxEmailCaller.Disabled = False
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        End If

        If InStr (1, EditSubjectCase.Text, "Systemänderbarkeit_SAP", vbTextCompare) then
        CaseProblem.Disabled = False
        CaseProblem.Disabled = False
        ComboBoxEmailCaller.Disabled = False
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        End If

        If InStr (1, EditSubjectCase.Text, "#Prio 1 Incident# ", vbTextCompare) then
        CaseProblem.Disabled = False
        CaseProblem.Disabled = False
        ComboBoxEmailCaller.Disabled = False
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        End If
        If InStr (1, EditSubjectCase.Text, "Debugg_Modus_SAP", vbTextCompare) then
        CaseProblem.Disabled = False
        CaseProblem.Disabled = False
        ComboBoxEmailCaller.Disabled = False
        CaseDiagnosis.Disabled = False
        KeywordTree.Disabled = False
        Attachment.Disabled = False
        CaseAttributes.Disabled = False
        ComboIncidentStatus.Disabled = False
        End If
      
END SUB
SUB ButtonActionItemsAdd_Click()
        Dim textdata,texttemp
        If TextBoxActionItemsInput.Text= "" then
        MsgBox("Input value is missing.")
        Else
        texttemp = TextBoxActionItemsInput.Text
        textdata = hlObj.GetValue("IncidentAttribute.ActionItems",0,0,0,0)
        If Not textdata = "" Then
        textdata = textdata & CHR(10) & texttemp
        Else
        textdata = texttemp
        End If
        hlObj.SetValue "IncidentAttribute.ActionItems",0,0,0,textdata
        End If
      
END SUB
SUB ButtonActionItemsDel_Click()
        Dim delete
        delete = MsgBox("Delete all action items permanently?" ,4,"Delete Action Items")
        If delete = 6 Then
        hlObj.SetValue "IncidentAttribute.ActionItems",0,0,0,""
        End If
      
END SUB
