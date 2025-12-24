hlContext.EnableTrace
'Deaktiviern bzw. aktivieren aller Traces fuer ein Skript, Text = Logtext im App.Log
'Ermitteln der Locale ID fuer die Sprachauswahl.
'Selecting the Locale ID for the desired language.
Dim lcid
lcid = 0
lcid = hlContext.GetLocaleID
Dim LangID
LangID = 0
LangID = hlContext.LangIDFromLCID(lcid)

'Aktuelles Objekt ermitteln.
'Detect the current object.
Dim hlCase
Set hlCase = Nothing
Set hlCase = hlContext.GetCurrentObject

Dim Editor
Editor = hlCase.GetValue("SUINFO.EDITOR", 0, 0, 0, 0)
Dim ActDate
ActDate = cstr(now)


'VB-Dictionary anlegen.
'Create VB-Dictionary.
Dim pCase
Set pCase = Nothing
Set pCase = CreateObject("Scripting.Dictionary")
pCase.CompareMode = vbTextCompare
pCase("BillCase") = False
pCase("attrOperation") = "IncidentSUAttribute.IncidentOperation"
pCase("attrDistinguishMixed") = "IncidentAttribute.RequestType"
pCase("Delegated") = False

'Vorgangsstatus auslesen.
'Retrieve Case status.
Dim state
state = ""
state = hlCase.GetValue("CASEINFO.INTERNALSTATE", 0, 0, 0, 0)

'Zuordnung von Internalstate zu OrderRequest-Status
'Mapping of Internalstate to OrderRequest-Status
Dim strOrdReqStatus
strOrdReqStatus = hlCase.GetValue("OrderRequestAttribute.OrderRequestStatus", 0, 0, 0, 0)

IF strOrdReqStatus = "OrderRequestStatusNew" Or strOrdReqStatus = "OrderRequestStatusOrdered" THEN
  hlCase.SetValue "CASEINFO.INTERNALSTATE", 0, 0, 0, "OPEN"
END IF
IF strOrdReqStatus = "OrderRequestStatusChangedStorno" Or strOrdReqStatus = "OrderRequestStatusStornoDelivered" Or strOrdReqStatus = "OrderRequestStatusDelivered" THEN
  hlCase.SetValue "CASEINFO.INTERNALSTATE", 0, 0, 0, "TOBECHECKED"
END IF
'If strOrdReqStatus = "OrderRequestStatusDelivered" Then
'	hlCase.SetValue"CASEINFO.INTERNALSTATE",0,0,0,"SOLVED"
'End If
IF strOrdReqStatus = "OrderRequestStatusClosure" THEN
  hlCase.SetValue "CASEINFO.INTERNALSTATE", 0, 0, 0, "CLOSED"
END IF


'Anfrager der letzten SU ermitteln.
'Retrieve the requester from the last SU.
Dim hlCaller
Set hlCaller = Nothing
Call hlITIL2.GetCallerLastSU(hlCase, hlCaller, hlContext)
IF hlITIL2.IsHLObject(hlCaller, hlContext) = True THEN
  Call hlITIL2.SetCaseInformation(hlCaller, hlCase, hlContext)
END IF


'Multiples Attribut Bestellpositionen abfragen und ggf. CI's anlegen
Dim OrderPosIDs
Set OrderPosIDs = Nothing
Dim PosID
PosID = 0
Dim CreateCI
CreateCI = 0
Dim Counter
Counter = 0
Dim CIisCreated
CIisCreated = 0
Dim CIType
CIType = 0
Dim CIQuantity
CIQuantity = 1
Dim CIQuantityInternal
CIQuantityInternal = 0
Dim ChangedOrderQuantity
ChangedOrderQuantity = 0
Dim i
i = 0
Dim NewCI
Set NewCI = Nothing
Dim Testname
Testname = ""
Dim OrderNumber
OrderNumber = 0
Dim CompanyCode
CompanyCode = 0
Dim OrderDate
OrderDate = 0
Dim OrderPosNr
OrderPosNr = 0
Dim VendorNumber
VendorNumber = 0
Dim VendorName
VendorName = 0
Dim AllocationNumber
AllocationNumber = 0
Dim AllocationType
AllocationType = ""
Dim PlaceOfUnloading
PlaceOfUnloading = ""
Dim Incorporation
Incorporation = ""
Dim PosOrderText
PosOrderText = ""
Dim Reciever
Reciever = ""
Dim cn
Set cn = Nothing
Dim rs
Set rs = Nothing
Dim CINumber
CINumber = ""
Dim QryString
QryString = ""
Dim Qry
Set Qry = Nothing
Dim AssetGroups
AssetGroups = ""
Dim AssetGroup
AssetGroup = ""
Dim AssetGroupID
AssetGroupID = ""
Dim Group
Set Group = Nothing
Dim ArticleDescription
ArticleDescription = ""
Dim CIPrice
CIPrice = 1
Dim CIPriceUnit
CIPriceUnit = "1"
Dim CIPriceCurrency
CIPriceCurrency = ""
Dim OrderText
OrderText = ""
Dim PosOrderInfoText
PosOrderInfoText = ""
Dim CIComment
CIComment = ""
Dim DeliveryDate
DeliveryDate = ""

OrderPosIDs = hlcase.GetContentIDs("OrderRequestAttribute.OrderedCIs_CA", 0)

'Allgemeingueltige Werte fuer alle CI's auslesen
OrderNumber = hlcase.GetValue("OrderRequestAttribute.OrderNumber", 0, 0, 0, 0)
CompanyCode = hlcase.GetValue("OrderRequestAttribute.CompanyCode", 0, 0, 0, 0)
VendorNumber = hlcase.GetValue("OrderRequestAttribute.VendorNumber", 0, 0, 0, 0)
VendorName = hlcase.GetValue("OrderRequestAttribute.VendorName", 0, 0, 0, 0)
OrderDate = hlcase.GetValue("OrderRequestAttribute.OrderDate", 0, 0, 0, 0)
OrderText = hlcase.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0)

'hlContext. Trace 1, "Gleich kommt die For each Schleife"

For Each PosID In OrderPosIDs
  'Counter = Counter + 1
  'hlContext. Trace 1, "Jetzt For each Schleife"
  'Pruefen ob CI erzeugt werden soll
  CreateCI = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CreateCI", 0, PosID, 0, 0)
  'Pruefen ob CI bereits erzeugt wurde
  CIisCreated = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CIisCreated", 0, PosID, 0, 0)
  IF CreateCI = "1" And CIisCreated <> "1" And strOrdReqStatus = "OrderRequestStatusOrdered" THEN
    'CI-Typ ermitteln
    CIType = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CIType", 0, PosID, 0, 0)
    'Anzahl der zu erstellenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    CIQuantityInternal = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    'Preiseinheit abfragen
    CIPriceUnit = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PriceUnit", 0, PosiD, 0, 0)
    'Bestellmengenaenderung abfragen
    ChangedOrderQuantity = hlITIL2.CheckIntegerValue(hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, 0), hlContext)
    'hlContext. Trace 1, ChangedOrderQuantity
    IF ChangedOrderQuantity > 0 THEN
      CIQuantity = ChangedOrderQuantity
    ELSE
      CIQuantity = CIQuantity
    END IF
    'Bestellposition
    OrderPosNr = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.OrderPosition", 0, PosID, 0, 0)
    'Abladestelle
    PlaceOfUnloading = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PlaceOfUnloading", 0, PosID, 0, 0)
    'Warenempfaenger
    Reciever = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.Reciever", 0, PosID, 0, 0)
    'Kontierungsnummer
    AllocationNumber = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.AllocationNumber", 0, PosID, 0, 0)
    'LieferDatum
    DeliveryDate = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.DeliveryDate", 0, PosID, 0, 0)
    'Kontierungstyp
    AllocationType = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.AllocationType", 0, PosID, 0, 0)
    'Positionsbestelltext
    PosOrderText = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PositionOrderText", 0, PosID, 0, 0)
    IF PosOrderText = "" THEN
      PosOrderText = " "
    END IF
    'Positionsinfotext
    PosOrderInfoText = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PositionInfoNotice", 0, PosID, 0, 0)
    IF PosOrderInfoText = "" THEN
      PosOrderInfoText = " "
    END IF
    'Werk/Standort
    Incorporation = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.Incorporation", 0, PosID, 0, 0)
    'Artikelbeschreibung
    ArticleDescription = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.ArticleDescription", 0, PosID, 0, 0)
    'Preis
    CIPrice = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_VALUE", 0, PosID, 0, 1)
    'Preiseinheit
    CIPriceCurrency = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_SYMBOL", 0, PosID, 0, 0)
    IF CIPriceUnit > 1 THEN
      CIPrice = clng(CIPrice) / clng(CIPriceUnit)
    ELSE
      CIPrice = CIPrice
    END IF
    CIComment = "Bestelltext/Ordertext: " & OrderText & CHR(13) & CHR(10) & CHR(13) & CHR(10)
    CIComment = CIComment & "Positionstext/Positiontext: " & PosOrderText & CHR(13) & CHR(10) & CHR(13) & CHR(10)
    CIComment = CIComment & "Positions-Infonotiz/Position-Infonotice: " & PosOrderInfoText
    SELECT CASE CIType
      'Arbeitsplatzcomputer/Desktopcomputer

      CASE "CITypeDesktopcomputer"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Computer anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("DesktopComputer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT desktop FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET desktop = desktop+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "DT00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "DT0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "DT000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "DT00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "DT0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "DT" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          'Neues CI dem Vorgang assoziieren
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Notebook
      CASE "CITypeNotebook"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("NotebookComputer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT notebook FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET notebook = notebook+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "NB00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "NB0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "NB000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "NB00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "NB0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "NB" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Software-Lizenz
      CASE "CITypeSoftware"
        'For i=1 To CIQuantity
        IF clng(CIQuantity) > 1 THEN
          CIPrice = clng(CIPrice) * clng(CIQuantity)
        END IF
        Set NewCI = hlContext.createobject("SoftwareLicense")
        NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
        NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
        NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
        NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
        NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
        NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
        NewCI.SetValue "SoftwareLicenseStatus.DocumentOrdered", 0, 0, 0, "1"
        NewCI.SetValue "TrumpfSoftwareStatus.SWPlannedAgent", 0, 0, 0, "System"
        NewCI.SetValue "TrumpfSoftwareStatus.SWPlannedDate", 0, 0, 0, ActDate
        NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
        NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
        NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
        NewCI.SetValue "TrumpfSoftwareStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
        NewCI.SetValue "SoftwareLicenseGeneral.SoftwareLicenseName", 0, 0, 0, ArticleDescription
        NewCI.SetValue "SoftwareLicenseCounter.ReferenceLicenseCount", 0, 0, 0, CIQuantity
        NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
        NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
        NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
        NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
        NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
        IF AllocationType = "K" THEN
          NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
        END IF
        IF AllocationType = "A" THEN
          NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
        END IF
        '------------------------------------------------------------------------------------------------
        'Generiert eine neue CI-Nummer
        Set cn = createobject("ADODB.Connection")

        'Verbindung oeffnen
        'Hier muss Server- und Datenbankname fest eingetragen werden!
        'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
        cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm1"
        cn.ConnectionTimeout = 10
        cn.Open

        'CI-Nummer auslesen
        Set rs = createobject("ADODB.Recordset")
        Set rs = cn.Execute("SELECT softwarelic FROM _cinumbers")
        'In Variable schreiben
        CINumber = rs.fields(0).value
        'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
        cn.execute("UPDATE _cinumbers SET softwarelic = softwarelic+1")

        'Verbindung schliessen
        rs.close
        cn.close
        IF LEN(cstr(CINumber)) = 1 THEN
          CINumber = "LI00000" + cstr(CINumber)
        END IF
        IF LEN(cstr(CINumber)) = 2 THEN
          CINumber = "LI0000" + cstr(CINumber)
        END IF
        IF LEN(cstr(CINumber)) = 3 THEN
          CINumber = "LI000" + cstr(CINumber)
        END IF
        IF LEN(cstr(CINumber)) = 4 THEN
          CINumber = "LI00" + cstr(CINumber)
        END IF
        IF LEN(cstr(CINumber)) = 5 THEN
          CINumber = "LI0" + cstr(CINumber)
        END IF
        IF LEN(cstr(CINumber)) = 6 THEN
          CINumber = "LI" + cstr(CINumber)
        END IF

        'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
        NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
        hlContext.saveobject NewCI
        Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
        'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
        'Zunaechst ID der Inventargruppe ermitteln
        QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
        'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
        Set Qry = hlContext.OpenSearch(QryString)
        IF Qry.GetItemCount(0, 0) = "1" THEN
          AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
          For Each Group In AssetGroups
            Set AssetGroup = Group
          Next
          Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
        END IF
        'Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Drucker
      CASE "CITypePrinter"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("Printer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "PrintSanDeviceDetail.PrintScanDeviceType", 0, 0, 0, "PSDTypePrinter"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT printer FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET printer = printer+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "PR00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "PR0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "PR000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "PR00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "PR0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "PR" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Kopierer
      CASE "CITypeCopyDevice"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("Printer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "PrintSanDeviceDetail.PrintScanDeviceType", 0, 0, 0, "PSDTypeCopyDevice"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT copydevice FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET copydevice = copydevice+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "CR00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "CR0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "CR000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "CR00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "CR0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "CR" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Multifunktionsgeraet
      CASE "CITypeMultifunctionDevice"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("Printer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "PrintSanDeviceDetail.PrintScanDeviceType", 0, 0, 0, "PSDTypeMultiFunctionDevice"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT multifunctiondevice FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET multifunctiondevice = multifunctiondevice+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "MF00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "MF0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "MF000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "MF00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "MF0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "MF" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Scanner
      CASE "CITypeScanner"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("Printer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "PrintSanDeviceDetail.PrintScanDeviceType", 0, 0, 0, "PSDTypeScanner"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT scanner FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET scanner = scanner+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "SC00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "SC0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "SC000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "SC00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "SC0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "SC" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Handy
      CASE "CITypeMobilePhone"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MobileDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "MobileDeviceDetail.MobileDeviceType", 0, 0, 0, "MobileDeviceTypeMobilePhone"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT handys FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET handys = handys+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "MP00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "MP0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "MP000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "MP00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "MP0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "MP" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'SIM-Karte
      CASE "CITypeSIMCard"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MobileDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "MobileDeviceDetail.MobileDeviceType", 0, 0, 0, "MobileDeviceTypeSIMCard"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT simcard FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET simcard = simcard+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "SI00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "SI0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "SI000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "SI00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "SI0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "SI" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'UMTS-Karte
      CASE "CITypeUMTSCard"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MobileDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "MobileDeviceDetail.MobileDeviceType", 0, 0, 0, "MobileDeviceTypeUMTSCard"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT umtscard FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET umtscard = umtscard+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "UM00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "UM0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "UM000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "UM00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "UM0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "UM" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'PDA
      CASE "CITypePDA"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MobileDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "MobileDeviceDetail.MobileDeviceType", 0, 0, 0, "MobileDeviceTypePDA"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT pda FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET pda = pda+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "PD00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "PD0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "PD000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "PD00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "PD0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "PD" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'BlackBerry
      CASE "CITypeBlackberry"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MobileDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "MobileDeviceDetail.MobileDeviceType", 0, 0, 0, "MobileDeviceTypeBlackBerry"
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT blackberry FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET blackberry = blackberry+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "BB00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "BB0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "BB000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "BB00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "BB0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "BB" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Monitor
      CASE "CITypeMonitor"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("Monitor")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT monitor FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET monitor = monitor+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "MO00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "MO0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "MO000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "MO00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "MO0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "MO" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Beamer
      CASE "CITypeBeamer"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MultiMediaDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "MultiMediaDeviceDetail.MultiMediaDeviceType", 0, 0, 0, "MultiMediaDeviceTypeBeamer"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT beamer FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET beamer = beamer+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "VP00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "VP0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "VP000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "VP00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "VP0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "VP" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Videokonferenztechnik
      CASE "CITypeVideoconferencetechnic"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MultiMediaDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "MultiMediaDeviceDetail.MultiMediaDeviceType", 0, 0, 0, "MultiMediaDeviceTypeVideoConferenceTechnic"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT videoconference FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET videoconference = videoconference+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "VC00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "VC0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "VC000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "VC00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "VC0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "VC" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Medientechnik
      CASE "CITypeMediaTechnic"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("MultiMediaDevice")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "MultiMediaDeviceDetail.MultiMediaDeviceType", 0, 0, 0, "MultiMediaDeviceTypeMediaTechnic"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT mediatechnic FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET mediatechnic = mediatechnic+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "MU00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "MU0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "MU000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "MU00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "MU0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "MU" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Diktiergeraet
      CASE "CITypeDictaphone"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeDictationDevice"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT diktiersystem FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET diktiersystem = diktiersystem+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "DS00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "DS0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "DS000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "DS00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "DS0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "DS" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'USV
      CASE "CITypeUSV"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeUSV"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT usv FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET usv = usv+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "UP00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "UP0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "UP000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "UP00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "UP0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "UP" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'ueberwachungskamera
      CASE "CITypeControlCam"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeControlCam"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT controlcam FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET controlcam = controlcam+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "MC00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "MC0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "MC000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "MC00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "MC0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "MC" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'BDE
      CASE "CITypeBDE"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeBDE"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT bde FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET bde = bde+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "DA00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "DA0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "DA000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "DA00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "DA0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "DA" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Spacemaus
      CASE "CITypeSpacemouse"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeSpaceMouse"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT spacemouse FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET spacemouse = spacemouse+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "SP00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "SP0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "SP000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "SP00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "SP0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "SP" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Aktive Netzwerkkomponente
      CASE "CITypeNetworkComponent"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("NetworkComponent")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "NetworkComponentDetail.NetworkComponentType", 0, 0, 0, "TypeActiveNetworkComponet"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT networkcomponent FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET networkcomponent = networkcomponent+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "AN00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "AN0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "AN000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "AN00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "AN0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "AN" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
      CASE "CITypeHomeOfficeRouter"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("NetworkComponent")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "NetworkComponentDetail.NetworkComponentType", 0, 0, 0, "TypeHomeOfficeRouter"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT homeofficerouter FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET homeofficerouter = homeofficerouter+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "HO00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "HO0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "HO000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "HO00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "HO0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "HO" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
        'Headset
      CASE "CITypeHeadset"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeHeadset"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT headset FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET headset = headset+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "HS00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "HS0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "HS000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "HS00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "HS0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "HS" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"

        'ConferencePhone
      CASE "CITypeConferencePhone"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Notebook anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("GenericAsset")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "GenericAssetDetail.GenericAssetType", 0, 0, 0, "GenericAssetTypeConferencePhone"
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm1"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT conferencephone FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET conferencephone = conferencephone+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "CP00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "CP0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "CP000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "CP00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "CP0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "CP" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          'hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"

        'Server
      CASE "CITypeServerComputer"
        For i = 1 To CIQuantity
          'hlContext.Trace 1, "Computer anlegen Nummer: " & i
          Set NewCI = hlContext.createobject("ServerComputer")
          NewCI.SetValue "ProcurementDetail.OrderNumber", 0, 0, 0, OrderNumber
          NewCI.SetValue "TrumpfAssetGeneral.VendorName", 0, 0, 0, VendorName
          NewCI.SetValue "ProcurementDetail.OrderDate", 0, 0, 0, OrderDate
          NewCI.SetValue "TrumpfAssetGeneral.CompanyCode", 0, 0, 0, CompanyCode
          NewCI.SetValue "ProcurementDetail.VendorNumber", 0, 0, 0, VendorNumber
          NewCI.SetValue "ProcurementDetail.OrderPosition", 0, 0, 0, OrderPosNr
          NewCI.SetValue "ProcurementDetail.AllocationNumber", 0, 0, 0, AllocationNumber
          NewCI.SetValue "ProcurementDetail.AllocationType", 0, 0, 0, AllocationType
          NewCI.SetValue "TrumpfAssetGeneral.GoodsRecipient", 0, 0, 0, Reciever
          NewCI.SetValue "TrumpfAssetGeneral.OrderPosID", 0, 0, 0, PosID
          NewCI.SetValue "TrumpfAssetGeneral.PlaceOfUnloading", 0, 0, 0, PlaceOfUnloading
          NewCI.SetValue "AssetGeneral.LongComment", 0, 0, 0, CIComment
          NewCI.SetValue "TrumpfAssetStatus.CIOrder", 0, 0, 0, "1"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderAgent", 0, 0, 0, "System"
          NewCI.SetValue "TrumpfAssetStatus.CIOrderDate", 0, 0, 0, ActDate
          NewCI.SetValue "ProcurementDetail.DeliveryDate", 0, 0, 0, DeliveryDate
          NewCI.SetValue "AssetGeneral.AssetName", 0, 0, 0, ArticleDescription
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_VALUE", 0, 0, 0, CIPrice
          NewCI.SetValue "AccountingDetail.PurchasePrice.CURRENCY_SYMBOL", 0, 0, 0, CIPriceCurrency
          NewCI.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, "Beschaffung/Order am/at: " + ActDate + " von/by: System" & vbNewLine
          IF AllocationType = "K" THEN
            NewCI.SetValue "AccountingDetail.CostCenter", 0, 0, 0, AllocationNumber
          END IF
          IF AllocationType = "A" THEN
            NewCI.SetValue "TrumpfAssetGeneral.InvestmentNumber", 0, 0, 0, AllocationNumber
          END IF
          '------------------------------------------------------------------------------------------------
          'Generiert eine neue CI-Nummer
          Set cn = createobject("ADODB.Connection")

          'Verbindung oeffnen
          'Hier muss Server- und Datenbankname fest eingetragen werden!
          'Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
          cn.ConnectionString = "Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2"
          cn.ConnectionTimeout = 10
          cn.Open

          'CI-Nummer auslesen
          Set rs = createobject("ADODB.Recordset")
          Set rs = cn.Execute("SELECT server FROM _cinumbers")
          'In Variable schreiben
          CINumber = rs.fields(0).value
          'CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
          cn.execute("UPDATE _cinumbers SET server = server+1")

          'Verbindung schliessen
          rs.close
          cn.close
          IF LEN(cstr(CINumber)) = 1 THEN
            CINumber = "SR00000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 2 THEN
            CINumber = "SR0000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 3 THEN
            CINumber = "SR000" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 4 THEN
            CINumber = "SR00" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 5 THEN
            CINumber = "SR0" + cstr(CINumber)
          END IF
          IF LEN(cstr(CINumber)) = 6 THEN
            CINumber = "SR" + cstr(CINumber)
          END IF

          'hlContext.Trace 1, "Ticket-ID = " & NextTicketID
          NewCI.SetValue "TrumpfAssetGeneral.CINumber", 0, 0, 0, CINumber
          hlContext.saveobject NewCI
          'Neues CI dem Vorgang assoziieren
          Call hlContext.CreateAssociation(hlcase, NewCI, 119155)
          'Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
          'Zunaechst ID der Inventargruppe ermitteln
          QryString = "SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = " & Incorporation
          hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
          Set Qry = hlContext.OpenSearch(QryString)
          IF Qry.GetItemCount(0, 0) = "1" THEN
            AssetGroups = Qry.GetItems(0, - 1, - 1, 0)
            For Each Group In AssetGroups
              Set AssetGroup = Group
            Next
            Call hlContext.CreateAssociation(AssetGroup, NewCI, 100706)
          END IF
        Next
        'Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
        'CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
        'hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
        'hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
        'Geaenderte Bestellmenge auf 0 setzen
        hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity", 0, PosID, 0, "0"
    END SELECT
    'Kennzeichnen, dass CI erzeugt wurde
    hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.CIisCreated", 0, PosID, 0, "1"
  END IF
  CreateCI = 0
Next

'Eliminierung von Geraeten------------------------------------------------------------------------------------------------------
'Pruefen, ob Anzahl der assoziierten CI's pro Typ groesser ist, als Inhalt des Attributs Bestellmenge
'Wenn ja, dann entsprechend viele CI's (Differenz aus Anzahl und Bestellmenge) auf Status "eliminiert" setzen
'Das Ganze nur bei OrderStatus Aenderung/Storno
Dim objs
Set objs = Nothing
Dim obj
Set obj = Nothing
Dim objtype
Set objtype = Nothing
Dim cistatus
cistatus = ""
Dim statuscounter
statuscounter = 0
Dim typecounter
typecounter = 0
Dim stornoquantity
stornoquantity = 0
Dim stornocounter
stornocounter = 0
Dim statusoverview
statusoverview = ""
Dim CIExistingAtSAPAM
CIExistingAtSAPAM = 0
Dim OrderPosID
OrderPosID = 0
Dim PosType
PosType = ""
For Each PosID In OrderPosIDs
  'Pruefen ob CI erzeugt werden soll
  CreateCI = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CreateCI", 0, PosID, 0, 0)
  'Pruefen ob CI bereits erzeugt wurde
  CIisCreated = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CIisCreated", 0, PosID, 0, 0)
  'Geraetetyp validieren
  PosType = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.CIType", 0, PosID, 0, 0)
  'Bestellmenge abfragen
  CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
  'Auf Eliminierung pruefen
  IF CreateCI = "1" And CIisCreated = "1" And strOrdReqStatus = "OrderRequestStatusChangedStorno" THEN
    'Anzahl assoziierte CIs ermitteln
    objs = hlcase.GetItems(0, - 1, - 1, 119155)
    'DesktopComputer
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "DesktopComputer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeDesktopcomputer" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "DesktopComputer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'NotebookComputer
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "NotebookComputer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeNotebook" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "NotebookComputer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Aktive Netzwerkomponente
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "NetworkComponent" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeNetworkComponent" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "NotebookComputer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Monitor
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "Monitor" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeMonitor" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "Monitor" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Printer
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "Printer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypePrinter" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "Printer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Scanner
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "Printer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeScanner" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "Printer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Kopierer
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "Printer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeCopyDevice" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "Printer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Multifunktionsgeraet
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "Printer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeMultifunctionDevice" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "Printer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Diktiergeraet
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeDictaphone" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Headset
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeHeadset" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'ConferencePhone
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeConferencePhone" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Spacemaus
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeSpacemouse" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'USV
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeUSV" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Ueberwachungskamera
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeControlCam" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'BDE
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "GenericAsset" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeBDE" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "GenericAsset" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Handy
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MobileDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeMobilePhone" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MobileDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'BlackBerry
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MobileDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeBlackberry" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MobileDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'PDA
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MobileDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypePDA" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MobileDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'SIM-Karte
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MobileDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeSIMCard" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MobileDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'UMTS-Karte
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MobileDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeUMTSCard" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MobileDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Videokonferenztechnik
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MultiMediaDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeVideoconferencetechnic" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MultiMediaDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Beamer
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MultiMediaDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeBeamer" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MultiMediaDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'Medientechnik
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "MultiMediaDevice" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeMediaTechnic" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "MultiMediaDevice" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
    'ServerComputer
    statuscounter = 0
    typecounter = 0
    stornocounter = 0
    For Each obj In objs
      objtype = obj.GetType()
      IF objtype = "ServerComputer" THEN
        OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
        IF clng(OrderPosID) = clng(PosID) THEN
          'Ist Geraet eliminiert?
          cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
          IF cistatus = "1" THEN
            statuscounter = statuscounter + 1
          END IF
          typecounter = typecounter + 1
        END IF
      END IF
    Next
    'Anzahl der zu eliminierenden CI`s ermitteln
    CIQuantity = hlcase.GetValue("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity", 0, PosID, 0, 0)
    stornoquantity = typecounter - CIQuantity - statuscounter
    IF stornoquantity > = "1" THEN
      'Jetzt die CIs eliminieren
      IF PosType = "CITypeServerComputer" THEN
        For Each obj In objs
          objtype = obj.GetType()
          IF objtype = "ServerComputer" THEN
            'OrderPosID des Geraets abfragen
            OrderPosID = obj.GetValue("TrumpfAssetGeneral.OrderPosID", 0, 0, 0, 0)
            IF clng(OrderPosID) = clng(PosID) THEN
              'Ist Geraet eliminiert?
              cistatus = obj.GetValue("TrumpfAssetStatus.CIElimination", 0, 0, 0, 0)
              IF cistatus <> "1" THEN
                obj.SetValue "TrumpfAssetStatus.CIElimination", 0, 0, 0, "1"
                obj.SetValue "TrumpfAssetStatus.CISubStatus", 0, 0, 0, "CISubStatusStorno"
                obj.SetValue "TrumpfAssetStatus.CIEliminationAgent", 0, 0, 0, "System"
                obj.SetValue "TrumpfAssetStatus.CIEliminationDate", 0, 0, 0, ActDate
                statusoverview = obj.GetValue("TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, 0)
                statusoverview = statusoverview & vbNewLine + "Eliminierung/Elimination am/at: " + ActDate + " durch/by: System" & vbNewLine
                obj.SetValue "TrumpfAssetStatus.CIStatusOverview", 0, 0, 0, statusoverview
                'Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                CIExistingAtSAPAM = obj.GetValue("TrumpfAssetGeneral.CIExistingAtSAPAM", 0, 0, 0, 0)
                IF CIExistingAtSAPAM = "1" THEN
                  Call ExportObjectIncident(hlContext, obj)
                  obj.SetValue "TrumpfAssetStatus.IncidentBecauseOfCIElimination", 0, 0, 0, "1"
                END IF
                hlContext.SaveObject obj
                stornocounter = stornocounter + 1
              END IF
            END IF
            IF Clng(stornocounter) = Clng(stornoquantity) THEN
              Exit For
            END IF
          END IF
        Next
      END IF
    END IF
  END IF
Next



'----------------------------------------------------------------------------------------------------------
'Beschreibung in das SU-Attribut CaseDescriptionSU kopieren.
'Copy Description to SU-Attribute CaseDescriptionSU.
'Die Indizes der SUs werden festgestellt
Dim suindices
suindices = hlcase.GetSvcUnitIndices()
Dim sumin
sumin = LBound(suindices) + 1
Dim sumindescr
sumindescr = ""
Dim Agent
Agent = ""
Dim Agent1
Agent1 = ""
Dim Last1SUIdx
Last1SUIdx = 0
Dim LastSU
LastSU = ""
'Index letzte SU
Last1SUIdx = hlITIL2.GetLastSUIdx(hlCase, hlContext)
'Index vorletzte SU
LastSU = Last1SUIdx - 1
Dim DescrText
DescrText = hlCase.GetValue("CaseDescription.DescriptionText", 0, 0, 0, 0)
'Urspruenlichen Beschreibungstext ermitteln
sumindescr = hlCase.GetValue("OrderRequestSUAttribute.CaseDescriptionSU", 0, 0, sumin, 0)
Agent = hlcase.GetValue("SUINFO.EDITOR", 0, 0, Last1SUIdx, 0)
Agent1 = hlcase.GetValue("SUINFO.EDITOR", 0, 0, sumin, 0)

'----------------------------------------------------------------------------------------------------------
'Kumuliert die Texte der Bearbeitungsschritte und schreibt sie in das
'Overview-Textfeld. Die Texte werden durch Trennzeichen voneinander abgegrenzt.
'Pruefen ob mehr als 1 SU
Dim DescrTextalt
DescrTextalt = hlCase.GetValue("OrderRequestSUAttribute.CaseDescriptionSU", 0, 0, LastSU, 0)
IF LastSU > 0 THEN
  'Pruefen, ob Beschreibungstext sich geaendert hat
  IF DescrText <> sumindescr THEN
    Agent = hlcase.GetValue("SUINFO.EDITOR", 0, 0, Last1SUIdx, 0)
    Dim DescriptionAll
    DescriptionAll = ""
    Dim ProblemAll
    ProblemAll = ""
    Dim ProblemAll1
    ProblemAll1 = ""
    Dim DiagnosisAll
    DiagnosisAll = ""
    Dim SolutionAll
    SolutionAll = ""
    Dim Problem, SUDiagnosis, SUActivity, SURegTime, Solution, Problemtitle, Diagnosistitle, Solutiontitle, ProblemtitleNew
    IF LangID = 7 THEN
      ProblemtitleNew = "=== Bestellbeschreibung neu ===" & " [von Agent : " & Agent & "]" & vbNewLine
      Problemtitle = "=== Urspruengliche Bestellbeschreibung ===" & " [von Agent : " & Agent1 & "]" & vbNewLine
      Diagnosistitle = "=== Taetigkeitsbeschreibungen ===" & vbNewLine
      Solutiontitle = "=== Loesungsbeschreibung ===" & " [von Agent : " & Agent & "]" & vbNewLine & vbNewLine
    ELSE
      ProblemtitleNew = "=== Orderdescription new===" & " [by Agent : " & Agent & "]" & vbNewLine
      Problemtitle = "=== Original Orderdescription ===" & " [by Agent : " & Agent1 & "]" & vbNewLine
      Diagnosistitle = "=== Diagnosisactivities ===" & vbNewLine
      Solutiontitle = "=== Final solution ===" & " [by Agent : " & Agent & "]" & vbNewLine
    END IF
    'Problem-, Diagnose- und Loesungstext auslesen und zusammenfassen
    Problem = DescrText
    Problem = Replace(Problem, Chr(13) & Chr(10), " ")
    IF Problem <> "" THEN
      ProblemAll1 = ProblemtitleNew & Problem & vbNewLine & String(80, "-") & vbNewLine
    END IF
    IF sumindescr <> "" THEN
      ProblemAll = Problemtitle & sumindescr & vbNewLine & String(80, "-") & vbNewLine
    END IF

    Dim SUIdx
    For Each SUIdx In suindices
      SUDiagnosis = hlcase.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, SUIdx, 0)
      'SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
      SURegTime = hlcase.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, SUIdx, 0)
      Agent = hlcase.GetValue("SUINFO.EDITOR", 0, 0, SUIdx, 0)
      IF SUDiagnosis = "" THEN
        IF LangID = 7 THEN
          SUDiagnosis = "<keine Beschreibung>"
        ELSE
          SUDiagnosis = "<no description>"
        END IF
      END IF
      DiagnosisAll = DiagnosisAll & SUIdx & ". SU (" & Agent & ") -> [" & SURegTime & "]:" & vbNewLine & SUDiagnosis & vbNewLine & String(80, "-") & vbNewLine
    Next
    Solution = hlcase.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0)
    Solution = Replace(Solution, Chr(13) & Chr(10), " ")
    IF LTrim(RTrim(Solution)) <> "" THEN
      SolutionAll = SolutionAll & Solutiontitle & Solution
    END IF
    'Gesammelte Texte in das uebersicht-Textfeld schreiben
    DescriptionAll = ProblemAll & ProblemAll1 & Diagnosistitle & DiagnosisAll & SolutionAll
    hlcase.SetValue "CaseGeneral.Overview", 0, 0, 0, DescriptionAll
  END IF
END IF
IF sumindescr = DescrText THEN
  Agent = hlcase.GetValue("SUINFO.EDITOR", 0, 0, Last1SUIdx, 0)

  IF LangID = 7 THEN
    Problemtitle = "=== Bestellbeschreibung ===" & " [von Agent : " & Agent1 & "]" & vbNewLine
    Diagnosistitle = "=== Taetigkeitsbeschreibungen ===" & vbNewLine
    Solutiontitle = "=== Loesungsbeschreibung ===" & " [von Agent : " & Agent1 & "]" & vbNewLine & vbNewLine
  ELSE
    Problemtitle = "=== Orderdescription ===" & " [by Agent : " & Agent1 & "]" & vbNewLine
    Diagnosistitle = "=== Diagnosisactivities ===" & vbNewLine
    Solutiontitle = "=== Final solution ===" & " [by Agent : " & Agent1 & "]" & vbNewLine
  END IF
  'Problem-, Diagnose- und Loesungstext auslesen und zusammenfassen
  Problem = DescrText
  Problem = Replace(Problem, Chr(13) & Chr(10), " ")
  IF LTrim(RTrim(Problem)) <> "" THEN
    ProblemAll = Problemtitle & Problem & vbNewLine & String(80, "-") & vbNewLine
  END IF

  '    Dim SUIdx
  For Each SUIdx In suindices
    SUDiagnosis = hlcase.GetValue("CaseDiagnosis.DiagnosisText", 0, 0, SUIdx, 0)
    'SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
    SURegTime = hlcase.GetValue("SUINFO.REGISTRATIONTIME", 0, 0, SUIdx, 0)
    Agent = hlcase.GetValue("SUINFO.EDITOR", 0, 0, SUIdx, 0)
    IF SUDiagnosis = "" THEN
      IF LangID = 7 THEN
        SUDiagnosis = "<keine Beschreibung>"
      ELSE
        SUDiagnosis = "<no description>"
      END IF
    END IF
    DiagnosisAll = DiagnosisAll & SUIdx & ". SU (" & Agent & ") -> [" & SURegTime & "]:" & vbNewLine & SUDiagnosis & vbNewLine & String(80, "-") & vbNewLine
  Next
  Solution = hlcase.GetValue("CaseSolution.SolutionText", 0, 0, 0, 0)
  Solution = Replace(Solution, Chr(13) & Chr(10), " ")
  IF LTrim(RTrim(Solution)) <> "" THEN
    SolutionAll = SolutionAll & Solutiontitle & Solution
  END IF
  'Gesammelte Texte in das uebersicht-Textfeld schreiben
  DescriptionAll = ProblemAll & Diagnosistitle & DiagnosisAll & SolutionAll
  hlcase.SetValue "CaseGeneral.Overview", 0, 0, 0, DescriptionAll
END IF

