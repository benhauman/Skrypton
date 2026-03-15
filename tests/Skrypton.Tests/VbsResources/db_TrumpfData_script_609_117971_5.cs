using System;
using System.Collections;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Compat;

namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }
        protected override GlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) => new GlobalReferences(compatLayer, env);
        protected override void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences globalReferences)
        {
            var _env = env ?? throw new ArgumentNullException(nameof(env));
            var _outer = globalReferences ?? throw new ArgumentNullException(nameof(globalReferences));

            _.CALLm1v0(this, _env.hlContext, "EnableTrace");
            //Deaktiviern bzw. aktivieren aller Traces fuer ein Skript, Text = Logtext im App.Log
            //Ermitteln der Locale ID fuer die Sprachauswahl.
            //Selecting the Locale ID for the desired language.
            _outer.lcid = (Int16)0;
            _outer.lcid = _.VAL(_.CALLm1v0(this, _env.hlContext, "GetLocaleID"));
            _outer.LangID = (Int16)0;
            _outer.LangID = _.VAL(_.CALLm1argp(this, _env.hlContext, "LangIDFromLCID", _.ARGS.Ref(_outer.lcid, v => { _outer.lcid = v; })));

            //Aktuelles Objekt ermitteln.
            //Detect the current object.
            _outer.hlCase = VBScriptConstants.Nothing;
            _outer.hlCase = _.OBJ(_.CALLm1v0(this, _env.hlContext, "GetCurrentObject"));

            _outer.Editor = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.ActDate = _.CSTR(_.NOW());

            //VB-Dictionary anlegen.
            //Create VB-Dictionary.
            _outer.pCase = VBScriptConstants.Nothing;
            _outer.pCase = _.OBJ(_.CREATEOBJECT("Scripting.Dictionary"));
            _.SET(VBScriptConstants.vbTextCompare, this, _outer.pCase, "CompareMode");
            _.SET(false, this, _outer.pCase, null, _.ARGS.Val("BillCase"));
            _.SET("IncidentSUAttribute.IncidentOperation", this, _outer.pCase, null, _.ARGS.Val("attrOperation"));
            _.SET("IncidentAttribute.RequestType", this, _outer.pCase, null, _.ARGS.Val("attrDistinguishMixed"));
            _.SET(false, this, _outer.pCase, null, _.ARGS.Val("Delegated"));

            //Vorgangsstatus auslesen.
            //Retrieve Case status.
            _outer.state = "";
            _outer.state = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CASEINFO.INTERNALSTATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            //Zuordnung von Internalstate zu OrderRequest-Status
            //Mapping of Internalstate to OrderRequest-Status
            _outer.strOrdReqStatus = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderRequestStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            if (_.IF(_.OR(_.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusNew"), _.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusOrdered"))))
            {
                _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CASEINFO.INTERNALSTATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("OPEN"));
            }
            if (_.IF(_.OR(_.OR(_.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusChangedStorno"), _.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusStornoDelivered")), _.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusDelivered"))))
            {
                _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CASEINFO.INTERNALSTATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("TOBECHECKED"));
            }
            //If strOrdReqStatus = "OrderRequestStatusDelivered" Then
            //	hlCase.SetValue"CASEINFO.INTERNALSTATE",0,0,0,"SOLVED"
            //End If
            if (_.IF(_.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusClosure")))
            {
                _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CASEINFO.INTERNALSTATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CLOSED"));
            }

            //Anfrager der letzten SU ermitteln.
            //Retrieve the requester from the last SU.
            _outer.hlCaller = VBScriptConstants.Nothing;
            _.CALLm1argp(this, _env.hlITIL2, "GetCallerLastSU", _.ARGS.Ref(_outer.hlCase, v2 => { _outer.hlCase = v2; }).Ref(_outer.hlCaller, v3 => { _outer.hlCaller = v3; }).Ref(_env.hlContext, v4 => { _env.hlContext = v4; }));
            if (_.IF(_.EQ(_.CALLm1argp(this, _env.hlITIL2, "IsHLObject", _.ARGS.Ref(_outer.hlCaller, v5 => { _outer.hlCaller = v5; }).Ref(_env.hlContext, v6 => { _env.hlContext = v6; })), true)))
            {
                _.CALLm1argp(this, _env.hlITIL2, "SetCaseInformation", _.ARGS.Ref(_outer.hlCaller, v7 => { _outer.hlCaller = v7; }).Ref(_outer.hlCase, v8 => { _outer.hlCase = v8; }).Ref(_env.hlContext, v9 => { _env.hlContext = v9; }));
            }

            //Multiples Attribut Bestellpositionen abfragen und ggf. CI's anlegen
            _outer.OrderPosIDs = VBScriptConstants.Nothing;
            _outer.PosID = (Int16)0;
            _outer.CreateCI = (Int16)0;
            _outer.Counter = (Int16)0;
            _outer.CIisCreated = (Int16)0;
            _outer.CIType = (Int16)0;
            _outer.CIQuantity = (Int16)1;
            _outer.CIQuantityInternal = (Int16)0;
            _outer.ChangedOrderQuantity = (Int16)0;
            _outer.i = (Int16)0;
            _outer.NewCI = VBScriptConstants.Nothing;
            _outer.Testname = "";
            _outer.OrderNumber = (Int16)0;
            _outer.CompanyCode = (Int16)0;
            _outer.OrderDate = (Int16)0;
            _outer.OrderPosNr = (Int16)0;
            _outer.VendorNumber = (Int16)0;
            _outer.VendorName = (Int16)0;
            _outer.AllocationNumber = (Int16)0;
            _outer.AllocationType = "";
            _outer.PlaceOfUnloading = "";
            _outer.Incorporation = "";
            _outer.PosOrderText = "";
            _outer.Reciever = "";
            _outer.cn = VBScriptConstants.Nothing;
            _outer.rs = VBScriptConstants.Nothing;
            _outer.CINumber = "";
            _outer.QryString = "";
            _outer.Qry = VBScriptConstants.Nothing;
            _outer.AssetGroups = "";
            _outer.AssetGroup = "";
            _outer.AssetGroupID = "";
            _outer.rewritten_Group = VBScriptConstants.Nothing;
            _outer.ArticleDescription = "";
            _outer.CIPrice = (Int16)1;
            _outer.CIPriceUnit = "1";
            _outer.CIPriceCurrency = "";
            _outer.OrderText = "";
            _outer.PosOrderInfoText = "";
            _outer.CIComment = "";
            _outer.DeliveryDate = "";

            _outer.OrderPosIDs = _.VAL(_.CALLm1v2(this, _outer.hlCase, "GetContentIDs", "OrderRequestAttribute.OrderedCIs_CA", (Int16)0));

            //Allgemeingueltige Werte fuer alle CI's auslesen
            _outer.OrderNumber = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.CompanyCode = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.VendorNumber = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.VendorName = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.OrderDate = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            _outer.OrderText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            //hlContext. Trace 1, "Gleich kommt die For each Schleife"

            var enumerationContent = _.ENUMERABLE(_outer.OrderPosIDs).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                _outer.PosID = enumerationContent.Current;
                //Counter = Counter + 1
                //hlContext. Trace 1, "Jetzt For each Schleife"
                //Pruefen ob CI erzeugt werden soll
                _outer.CreateCI = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CreateCI").Val((Int16)0).Ref(_outer.PosID, v10 => { _outer.PosID = v10; }).Val((Int16)0).Val((Int16)0)));
                //Pruefen ob CI bereits erzeugt wurde
                _outer.CIisCreated = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIisCreated").Val((Int16)0).Ref(_outer.PosID, v11 => { _outer.PosID = v11; }).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.AND(_.AND(_.EQ(_.NullableSTR(_outer.CreateCI), "1"), _.NOTEQ(_.NullableSTR(_outer.CIisCreated), "1")), _.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusOrdered"))))
                {
                    //CI-Typ ermitteln
                    _outer.CIType = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIType").Val((Int16)0).Ref(_outer.PosID, v12 => { _outer.PosID = v12; }).Val((Int16)0).Val((Int16)0)));
                    //Anzahl der zu erstellenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v13 => { _outer.PosID = v13; }).Val((Int16)0).Val((Int16)0)));
                    _outer.CIQuantityInternal = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v14 => { _outer.PosID = v14; }).Val((Int16)0).Val((Int16)0)));
                    //Preiseinheit abfragen
                    _outer.CIPriceUnit = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PriceUnit").Val((Int16)0).Ref(_outer.PosID, v15 => { _outer.PosID = v15; }).Val((Int16)0).Val((Int16)0)));
                    //Bestellmengenaenderung abfragen
                    _outer.ChangedOrderQuantity = _.VAL(_.CALLm1v2(this, _env.hlITIL2, "CheckIntegerValue", _.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v16 => { _outer.PosID = v16; }).Val((Int16)0).Val((Int16)0)), _env.hlContext));
                    //hlContext. Trace 1, ChangedOrderQuantity
                    if (_.IF(_.GT(_.NullableNUM(_outer.ChangedOrderQuantity), (Int16)0)))
                    {
                        _outer.CIQuantity = _.VAL(_outer.ChangedOrderQuantity);
                    }
                    else
                    {
                        _outer.CIQuantity = _.VAL(_outer.CIQuantity);
                    }
                    //Bestellposition
                    _outer.OrderPosNr = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.OrderPosition").Val((Int16)0).Ref(_outer.PosID, v17 => { _outer.PosID = v17; }).Val((Int16)0).Val((Int16)0)));
                    //Abladestelle
                    _outer.PlaceOfUnloading = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PlaceOfUnloading").Val((Int16)0).Ref(_outer.PosID, v18 => { _outer.PosID = v18; }).Val((Int16)0).Val((Int16)0)));
                    //Warenempfaenger
                    _outer.Reciever = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.Reciever").Val((Int16)0).Ref(_outer.PosID, v19 => { _outer.PosID = v19; }).Val((Int16)0).Val((Int16)0)));
                    //Kontierungsnummer
                    _outer.AllocationNumber = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.AllocationNumber").Val((Int16)0).Ref(_outer.PosID, v20 => { _outer.PosID = v20; }).Val((Int16)0).Val((Int16)0)));
                    //LieferDatum
                    _outer.DeliveryDate = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.DeliveryDate").Val((Int16)0).Ref(_outer.PosID, v21 => { _outer.PosID = v21; }).Val((Int16)0).Val((Int16)0)));
                    //Kontierungstyp
                    _outer.AllocationType = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.AllocationType").Val((Int16)0).Ref(_outer.PosID, v22 => { _outer.PosID = v22; }).Val((Int16)0).Val((Int16)0)));
                    //Positionsbestelltext
                    _outer.PosOrderText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PositionOrderText").Val((Int16)0).Ref(_outer.PosID, v23 => { _outer.PosID = v23; }).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(_outer.PosOrderText), "")))
                    {
                        _outer.PosOrderText = " ";
                    }
                    //Positionsinfotext
                    _outer.PosOrderInfoText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PositionInfoNotice").Val((Int16)0).Ref(_outer.PosID, v24 => { _outer.PosID = v24; }).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(_outer.PosOrderInfoText), "")))
                    {
                        _outer.PosOrderInfoText = " ";
                    }
                    //Werk/Standort
                    _outer.Incorporation = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.Incorporation").Val((Int16)0).Ref(_outer.PosID, v25 => { _outer.PosID = v25; }).Val((Int16)0).Val((Int16)0)));
                    //Artikelbeschreibung
                    _outer.ArticleDescription = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ArticleDescription").Val((Int16)0).Ref(_outer.PosID, v26 => { _outer.PosID = v26; }).Val((Int16)0).Val((Int16)0)));
                    //Preis
                    _outer.CIPrice = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Ref(_outer.PosID, v27 => { _outer.PosID = v27; }).Val((Int16)0).Val((Int16)1)));
                    //Preiseinheit
                    _outer.CIPriceCurrency = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Ref(_outer.PosID, v28 => { _outer.PosID = v28; }).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.GT(_.NullableNUM(_outer.CIPriceUnit), (Int16)1)))
                    {
                        _outer.CIPrice = _.DIV(_.CLNG(_outer.CIPrice), _.CLNG(_outer.CIPriceUnit));
                    }
                    else
                    {
                        _outer.CIPrice = _.VAL(_outer.CIPrice);
                    }
                    _outer.CIComment = _.CONCAT("Bestelltext/Ordertext: ", _outer.OrderText, _.CHR((Int16)13), _.CHR((Int16)10), _.CHR((Int16)13), _.CHR((Int16)10));
                    _outer.CIComment = _.CONCAT(_outer.CIComment, "Positionstext/Positiontext: ", _outer.PosOrderText, _.CHR((Int16)13), _.CHR((Int16)10), _.CHR((Int16)13), _.CHR((Int16)10));
                    _outer.CIComment = _.CONCAT(_outer.CIComment, "Positions-Infonotiz/Position-Infonotice: ", _outer.PosOrderInfoText);
                    //Arbeitsplatzcomputer/Desktopcomputer
                    if (_.IF(_.EQ(_outer.CIType, "CITypeDesktopcomputer")))
                    {
                        var loopEnd = _.NUM(_outer.CIQuantity);
                        var loopStart = _.NUM((Int16)1, loopEnd);
                        if (_.StrictLTE(loopStart, loopEnd))
                        {
                            for (_outer.i = loopStart; _.StrictLTE(_outer.i, loopEnd); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Computer anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "DesktopComputer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v29 => { _outer.OrderNumber = v29; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v30 => { _outer.VendorName = v30; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v31 => { _outer.OrderDate = v31; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v32 => { _outer.CompanyCode = v32; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v33 => { _outer.VendorNumber = v33; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v34 => { _outer.OrderPosNr = v34; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v35 => { _outer.AllocationNumber = v35; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v36 => { _outer.AllocationType = v36; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v37 => { _outer.Reciever = v37; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v38 => { _outer.PosID = v38; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v39 => { _outer.PlaceOfUnloading = v39; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v40 => { _outer.CIComment = v40; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v41 => { _outer.ActDate = v41; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v42 => { _outer.DeliveryDate = v42; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v43 => { _outer.ArticleDescription = v43; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v44 => { _outer.CIPrice = v44; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v45 => { _outer.CIPriceCurrency = v45; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v46 => { _outer.AllocationNumber = v46; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v47 => { _outer.AllocationNumber = v47; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT desktop FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET desktop = desktop+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("DT00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("DT0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("DT000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("DT00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("DT0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("DT", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v48 => { _outer.CINumber = v48; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v49 => { _outer.NewCI = v49; }));
                                //Neues CI dem Vorgang assoziieren
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v50 => { _outer.hlCase = v50; }).Ref(_outer.NewCI, v51 => { _outer.NewCI = v51; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v52 => { _outer.QryString = v52; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent2 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent2.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent2.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v53 => { _outer.AssetGroup = v53; }).Ref(_outer.NewCI, v54 => { _outer.NewCI = v54; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v55 => { _outer.PosID = v55; }).Val((Int16)0).Val("0"));
                        //Notebook
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeNotebook")))
                    {
                        var loopEnd2 = _.NUM(_outer.CIQuantity);
                        var loopStart2 = _.NUM((Int16)1, loopEnd2);
                        if (_.StrictLTE(loopStart2, loopEnd2))
                        {
                            for (_outer.i = loopStart2; _.StrictLTE(_outer.i, loopEnd2); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "NotebookComputer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v56 => { _outer.OrderNumber = v56; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v57 => { _outer.VendorName = v57; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v58 => { _outer.OrderDate = v58; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v59 => { _outer.CompanyCode = v59; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v60 => { _outer.VendorNumber = v60; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v61 => { _outer.OrderPosNr = v61; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v62 => { _outer.AllocationNumber = v62; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v63 => { _outer.AllocationType = v63; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v64 => { _outer.Reciever = v64; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v65 => { _outer.PosID = v65; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v66 => { _outer.PlaceOfUnloading = v66; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v67 => { _outer.CIComment = v67; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v68 => { _outer.ActDate = v68; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v69 => { _outer.DeliveryDate = v69; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v70 => { _outer.ArticleDescription = v70; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v71 => { _outer.CIPrice = v71; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v72 => { _outer.CIPriceCurrency = v72; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v73 => { _outer.AllocationNumber = v73; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v74 => { _outer.AllocationNumber = v74; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT notebook FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET notebook = notebook+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("NB00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("NB0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("NB000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("NB00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("NB0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("NB", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v75 => { _outer.CINumber = v75; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v76 => { _outer.NewCI = v76; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v77 => { _outer.hlCase = v77; }).Ref(_outer.NewCI, v78 => { _outer.NewCI = v78; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v79 => { _outer.QryString = v79; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent3 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent3.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent3.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v80 => { _outer.AssetGroup = v80; }).Ref(_outer.NewCI, v81 => { _outer.NewCI = v81; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v82 => { _outer.PosID = v82; }).Val((Int16)0).Val("0"));
                        //Software-Lizenz
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeSoftware")))
                    {
                        //For i=1 To CIQuantity
                        if (_.IF(_.GT(_.NullableNUM(_.CLNG(_outer.CIQuantity)), (Int16)1)))
                        {
                            _outer.CIPrice = _.MULT(_.CLNG(_outer.CIPrice), _.CLNG(_outer.CIQuantity));
                        }
                        _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "SoftwareLicense"));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v83 => { _outer.OrderNumber = v83; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v84 => { _outer.OrderDate = v84; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v85 => { _outer.VendorNumber = v85; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v86 => { _outer.OrderPosNr = v86; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v87 => { _outer.AllocationNumber = v87; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v88 => { _outer.AllocationType = v88; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseStatus.DocumentOrdered").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.SWPlannedAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.SWPlannedDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v89 => { _outer.ActDate = v89; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v90 => { _outer.DeliveryDate = v90; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v91 => { _outer.CIPrice = v91; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v92 => { _outer.CIPriceCurrency = v92; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseGeneral.SoftwareLicenseName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v93 => { _outer.ArticleDescription = v93; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIQuantity, v94 => { _outer.CIQuantity = v94; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v95 => { _outer.VendorName = v95; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v96 => { _outer.CompanyCode = v96; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v97 => { _outer.Reciever = v97; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v98 => { _outer.PosID = v98; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v99 => { _outer.PlaceOfUnloading = v99; }));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                        {
                            _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v100 => { _outer.AllocationNumber = v100; }));
                        }
                        if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                        {
                            _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v101 => { _outer.AllocationNumber = v101; }));
                        }
                        //------------------------------------------------------------------------------------------------
                        //Generiert eine neue CI-Nummer
                        _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                        //Verbindung oeffnen
                        //Hier muss Server- und Datenbankname fest eingetragen werden!
                        //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                        _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm1", this, _outer.cn, "ConnectionString");
                        _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                        _.CALLm1v0(this, _outer.cn, "Open");

                        //CI-Nummer auslesen
                        _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                        _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT softwarelic FROM _cinumbers"));
                        //In Variable schreiben
                        _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                        //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                        _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET softwarelic = softwarelic+1");

                        //Verbindung schliessen
                        _.CALLm1v0(this, _outer.rs, "close");
                        _.CALLm1v0(this, _outer.cn, "close");
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                        {
                            _outer.CINumber = _.ADD("LI00000", _.CSTR(_outer.CINumber));
                        }
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                        {
                            _outer.CINumber = _.ADD("LI0000", _.CSTR(_outer.CINumber));
                        }
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                        {
                            _outer.CINumber = _.ADD("LI000", _.CSTR(_outer.CINumber));
                        }
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                        {
                            _outer.CINumber = _.ADD("LI00", _.CSTR(_outer.CINumber));
                        }
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                        {
                            _outer.CINumber = _.ADD("LI0", _.CSTR(_outer.CINumber));
                        }
                        if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                        {
                            _outer.CINumber = _.ADD("LI", _.CSTR(_outer.CINumber));
                        }

                        //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v102 => { _outer.CINumber = v102; }));
                        _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v103 => { _outer.NewCI = v103; }));
                        _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v104 => { _outer.hlCase = v104; }).Ref(_outer.NewCI, v105 => { _outer.NewCI = v105; }).Val(119155));
                        //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                        //Zunaechst ID der Inventargruppe ermitteln
                        _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                        //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                        _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v106 => { _outer.QryString = v106; })));
                        if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                        {
                            _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                            var enumerationContent4 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent4.MoveNext())
                                    break;
                                _outer.rewritten_Group = enumerationContent4.Current;
                                _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                            }
                            _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v107 => { _outer.AssetGroup = v107; }).Ref(_outer.NewCI, v108 => { _outer.NewCI = v108; }).Val(100706));
                        }
                        //Next
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v109 => { _outer.PosID = v109; }).Val((Int16)0).Val("0"));
                        //Drucker
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypePrinter")))
                    {
                        var loopEnd3 = _.NUM(_outer.CIQuantity);
                        var loopStart3 = _.NUM((Int16)1, loopEnd3);
                        if (_.StrictLTE(loopStart3, loopEnd3))
                        {
                            for (_outer.i = loopStart3; _.StrictLTE(_outer.i, loopEnd3); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "Printer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v110 => { _outer.OrderNumber = v110; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v111 => { _outer.VendorName = v111; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v112 => { _outer.OrderDate = v112; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v113 => { _outer.CompanyCode = v113; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v114 => { _outer.VendorNumber = v114; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v115 => { _outer.OrderPosNr = v115; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v116 => { _outer.AllocationNumber = v116; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v117 => { _outer.AllocationType = v117; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v118 => { _outer.Reciever = v118; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v119 => { _outer.PosID = v119; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v120 => { _outer.PlaceOfUnloading = v120; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v121 => { _outer.CIComment = v121; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypePrinter"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v122 => { _outer.ActDate = v122; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v123 => { _outer.DeliveryDate = v123; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v124 => { _outer.ArticleDescription = v124; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v125 => { _outer.CIPrice = v125; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v126 => { _outer.CIPriceCurrency = v126; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v127 => { _outer.AllocationNumber = v127; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v128 => { _outer.AllocationNumber = v128; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT printer FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET printer = printer+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("PR00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("PR0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("PR000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("PR00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("PR0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("PR", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v129 => { _outer.CINumber = v129; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v130 => { _outer.NewCI = v130; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v131 => { _outer.hlCase = v131; }).Ref(_outer.NewCI, v132 => { _outer.NewCI = v132; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v133 => { _outer.QryString = v133; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent5 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent5.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent5.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v134 => { _outer.AssetGroup = v134; }).Ref(_outer.NewCI, v135 => { _outer.NewCI = v135; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v136 => { _outer.PosID = v136; }).Val((Int16)0).Val("0"));
                        //Kopierer
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeCopyDevice")))
                    {
                        var loopEnd4 = _.NUM(_outer.CIQuantity);
                        var loopStart4 = _.NUM((Int16)1, loopEnd4);
                        if (_.StrictLTE(loopStart4, loopEnd4))
                        {
                            for (_outer.i = loopStart4; _.StrictLTE(_outer.i, loopEnd4); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "Printer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v137 => { _outer.OrderNumber = v137; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v138 => { _outer.VendorName = v138; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v139 => { _outer.OrderDate = v139; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v140 => { _outer.CompanyCode = v140; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v141 => { _outer.VendorNumber = v141; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v142 => { _outer.OrderPosNr = v142; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v143 => { _outer.AllocationNumber = v143; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v144 => { _outer.AllocationType = v144; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v145 => { _outer.Reciever = v145; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v146 => { _outer.PosID = v146; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v147 => { _outer.PlaceOfUnloading = v147; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v148 => { _outer.CIComment = v148; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeCopyDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v149 => { _outer.ActDate = v149; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v150 => { _outer.DeliveryDate = v150; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v151 => { _outer.ArticleDescription = v151; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v152 => { _outer.CIPrice = v152; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v153 => { _outer.CIPriceCurrency = v153; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v154 => { _outer.AllocationNumber = v154; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v155 => { _outer.AllocationNumber = v155; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT copydevice FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET copydevice = copydevice+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("CR00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("CR0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("CR000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("CR00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("CR0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("CR", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v156 => { _outer.CINumber = v156; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v157 => { _outer.NewCI = v157; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v158 => { _outer.hlCase = v158; }).Ref(_outer.NewCI, v159 => { _outer.NewCI = v159; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v160 => { _outer.QryString = v160; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent6 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent6.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent6.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v161 => { _outer.AssetGroup = v161; }).Ref(_outer.NewCI, v162 => { _outer.NewCI = v162; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v163 => { _outer.PosID = v163; }).Val((Int16)0).Val("0"));
                        //Multifunktionsgeraet
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeMultifunctionDevice")))
                    {
                        var loopEnd5 = _.NUM(_outer.CIQuantity);
                        var loopStart5 = _.NUM((Int16)1, loopEnd5);
                        if (_.StrictLTE(loopStart5, loopEnd5))
                        {
                            for (_outer.i = loopStart5; _.StrictLTE(_outer.i, loopEnd5); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "Printer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v164 => { _outer.OrderNumber = v164; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v165 => { _outer.VendorName = v165; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v166 => { _outer.OrderDate = v166; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v167 => { _outer.CompanyCode = v167; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v168 => { _outer.VendorNumber = v168; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v169 => { _outer.OrderPosNr = v169; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v170 => { _outer.AllocationNumber = v170; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v171 => { _outer.AllocationType = v171; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v172 => { _outer.Reciever = v172; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v173 => { _outer.PosID = v173; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v174 => { _outer.PlaceOfUnloading = v174; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v175 => { _outer.CIComment = v175; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeMultiFunctionDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v176 => { _outer.ActDate = v176; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v177 => { _outer.DeliveryDate = v177; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v178 => { _outer.ArticleDescription = v178; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v179 => { _outer.CIPrice = v179; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v180 => { _outer.CIPriceCurrency = v180; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v181 => { _outer.AllocationNumber = v181; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v182 => { _outer.AllocationNumber = v182; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT multifunctiondevice FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET multifunctiondevice = multifunctiondevice+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("MF00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("MF0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("MF000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("MF00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("MF0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("MF", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v183 => { _outer.CINumber = v183; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v184 => { _outer.NewCI = v184; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v185 => { _outer.hlCase = v185; }).Ref(_outer.NewCI, v186 => { _outer.NewCI = v186; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v187 => { _outer.QryString = v187; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent7 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent7.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent7.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v188 => { _outer.AssetGroup = v188; }).Ref(_outer.NewCI, v189 => { _outer.NewCI = v189; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v190 => { _outer.PosID = v190; }).Val((Int16)0).Val("0"));
                        //Scanner
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeScanner")))
                    {
                        var loopEnd6 = _.NUM(_outer.CIQuantity);
                        var loopStart6 = _.NUM((Int16)1, loopEnd6);
                        if (_.StrictLTE(loopStart6, loopEnd6))
                        {
                            for (_outer.i = loopStart6; _.StrictLTE(_outer.i, loopEnd6); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "Printer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v191 => { _outer.OrderNumber = v191; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v192 => { _outer.VendorName = v192; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v193 => { _outer.OrderDate = v193; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v194 => { _outer.CompanyCode = v194; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v195 => { _outer.VendorNumber = v195; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v196 => { _outer.OrderPosNr = v196; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v197 => { _outer.AllocationNumber = v197; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v198 => { _outer.AllocationType = v198; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v199 => { _outer.Reciever = v199; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v200 => { _outer.PosID = v200; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v201 => { _outer.PlaceOfUnloading = v201; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v202 => { _outer.CIComment = v202; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeScanner"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v203 => { _outer.ActDate = v203; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v204 => { _outer.DeliveryDate = v204; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v205 => { _outer.ArticleDescription = v205; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v206 => { _outer.CIPrice = v206; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v207 => { _outer.CIPriceCurrency = v207; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v208 => { _outer.AllocationNumber = v208; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v209 => { _outer.AllocationNumber = v209; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT scanner FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET scanner = scanner+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("SC00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("SC0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("SC000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("SC00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("SC0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("SC", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v210 => { _outer.CINumber = v210; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v211 => { _outer.NewCI = v211; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v212 => { _outer.hlCase = v212; }).Ref(_outer.NewCI, v213 => { _outer.NewCI = v213; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v214 => { _outer.QryString = v214; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent8 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent8.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent8.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v215 => { _outer.AssetGroup = v215; }).Ref(_outer.NewCI, v216 => { _outer.NewCI = v216; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v217 => { _outer.PosID = v217; }).Val((Int16)0).Val("0"));
                        //Handy
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeMobilePhone")))
                    {
                        var loopEnd7 = _.NUM(_outer.CIQuantity);
                        var loopStart7 = _.NUM((Int16)1, loopEnd7);
                        if (_.StrictLTE(loopStart7, loopEnd7))
                        {
                            for (_outer.i = loopStart7; _.StrictLTE(_outer.i, loopEnd7); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MobileDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v218 => { _outer.OrderNumber = v218; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v219 => { _outer.VendorName = v219; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v220 => { _outer.OrderDate = v220; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v221 => { _outer.CompanyCode = v221; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v222 => { _outer.VendorNumber = v222; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v223 => { _outer.OrderPosNr = v223; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v224 => { _outer.AllocationNumber = v224; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v225 => { _outer.AllocationType = v225; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v226 => { _outer.Reciever = v226; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v227 => { _outer.PosID = v227; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v228 => { _outer.PlaceOfUnloading = v228; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v229 => { _outer.CIComment = v229; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeMobilePhone"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v230 => { _outer.ActDate = v230; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v231 => { _outer.DeliveryDate = v231; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v232 => { _outer.ArticleDescription = v232; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v233 => { _outer.CIPrice = v233; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v234 => { _outer.CIPriceCurrency = v234; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v235 => { _outer.AllocationNumber = v235; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v236 => { _outer.AllocationNumber = v236; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT handys FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET handys = handys+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("MP00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("MP0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("MP000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("MP00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("MP0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("MP", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v237 => { _outer.CINumber = v237; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v238 => { _outer.NewCI = v238; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v239 => { _outer.hlCase = v239; }).Ref(_outer.NewCI, v240 => { _outer.NewCI = v240; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v241 => { _outer.QryString = v241; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent9 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent9.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent9.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v242 => { _outer.AssetGroup = v242; }).Ref(_outer.NewCI, v243 => { _outer.NewCI = v243; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v244 => { _outer.PosID = v244; }).Val((Int16)0).Val("0"));
                        //SIM-Karte
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeSIMCard")))
                    {
                        var loopEnd8 = _.NUM(_outer.CIQuantity);
                        var loopStart8 = _.NUM((Int16)1, loopEnd8);
                        if (_.StrictLTE(loopStart8, loopEnd8))
                        {
                            for (_outer.i = loopStart8; _.StrictLTE(_outer.i, loopEnd8); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MobileDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v245 => { _outer.OrderNumber = v245; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v246 => { _outer.VendorName = v246; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v247 => { _outer.OrderDate = v247; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v248 => { _outer.CompanyCode = v248; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v249 => { _outer.VendorNumber = v249; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v250 => { _outer.OrderPosNr = v250; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v251 => { _outer.AllocationNumber = v251; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v252 => { _outer.AllocationType = v252; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v253 => { _outer.Reciever = v253; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v254 => { _outer.PosID = v254; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v255 => { _outer.PlaceOfUnloading = v255; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v256 => { _outer.CIComment = v256; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeSIMCard"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v257 => { _outer.ActDate = v257; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v258 => { _outer.DeliveryDate = v258; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v259 => { _outer.ArticleDescription = v259; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v260 => { _outer.CIPrice = v260; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v261 => { _outer.CIPriceCurrency = v261; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v262 => { _outer.AllocationNumber = v262; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v263 => { _outer.AllocationNumber = v263; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT simcard FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET simcard = simcard+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("SI00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("SI0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("SI000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("SI00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("SI0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("SI", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v264 => { _outer.CINumber = v264; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v265 => { _outer.NewCI = v265; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v266 => { _outer.hlCase = v266; }).Ref(_outer.NewCI, v267 => { _outer.NewCI = v267; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v268 => { _outer.QryString = v268; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent10 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent10.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent10.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v269 => { _outer.AssetGroup = v269; }).Ref(_outer.NewCI, v270 => { _outer.NewCI = v270; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v271 => { _outer.PosID = v271; }).Val((Int16)0).Val("0"));
                        //UMTS-Karte
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeUMTSCard")))
                    {
                        var loopEnd9 = _.NUM(_outer.CIQuantity);
                        var loopStart9 = _.NUM((Int16)1, loopEnd9);
                        if (_.StrictLTE(loopStart9, loopEnd9))
                        {
                            for (_outer.i = loopStart9; _.StrictLTE(_outer.i, loopEnd9); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MobileDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v272 => { _outer.OrderNumber = v272; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v273 => { _outer.VendorName = v273; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v274 => { _outer.OrderDate = v274; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v275 => { _outer.CompanyCode = v275; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v276 => { _outer.VendorNumber = v276; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v277 => { _outer.OrderPosNr = v277; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v278 => { _outer.AllocationNumber = v278; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v279 => { _outer.AllocationType = v279; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v280 => { _outer.Reciever = v280; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v281 => { _outer.PosID = v281; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v282 => { _outer.PlaceOfUnloading = v282; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v283 => { _outer.CIComment = v283; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeUMTSCard"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v284 => { _outer.ActDate = v284; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v285 => { _outer.DeliveryDate = v285; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v286 => { _outer.ArticleDescription = v286; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v287 => { _outer.CIPrice = v287; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v288 => { _outer.CIPriceCurrency = v288; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v289 => { _outer.AllocationNumber = v289; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v290 => { _outer.AllocationNumber = v290; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT umtscard FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET umtscard = umtscard+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("UM00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("UM0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("UM000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("UM00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("UM0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("UM", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v291 => { _outer.CINumber = v291; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v292 => { _outer.NewCI = v292; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v293 => { _outer.hlCase = v293; }).Ref(_outer.NewCI, v294 => { _outer.NewCI = v294; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v295 => { _outer.QryString = v295; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent11 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent11.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent11.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v296 => { _outer.AssetGroup = v296; }).Ref(_outer.NewCI, v297 => { _outer.NewCI = v297; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v298 => { _outer.PosID = v298; }).Val((Int16)0).Val("0"));
                        //PDA
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypePDA")))
                    {
                        var loopEnd10 = _.NUM(_outer.CIQuantity);
                        var loopStart10 = _.NUM((Int16)1, loopEnd10);
                        if (_.StrictLTE(loopStart10, loopEnd10))
                        {
                            for (_outer.i = loopStart10; _.StrictLTE(_outer.i, loopEnd10); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MobileDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v299 => { _outer.OrderNumber = v299; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v300 => { _outer.VendorName = v300; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v301 => { _outer.OrderDate = v301; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v302 => { _outer.CompanyCode = v302; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v303 => { _outer.VendorNumber = v303; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v304 => { _outer.OrderPosNr = v304; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v305 => { _outer.AllocationNumber = v305; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v306 => { _outer.AllocationType = v306; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v307 => { _outer.Reciever = v307; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v308 => { _outer.PosID = v308; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v309 => { _outer.PlaceOfUnloading = v309; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v310 => { _outer.CIComment = v310; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypePDA"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v311 => { _outer.ActDate = v311; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v312 => { _outer.DeliveryDate = v312; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v313 => { _outer.ArticleDescription = v313; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v314 => { _outer.CIPrice = v314; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v315 => { _outer.CIPriceCurrency = v315; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v316 => { _outer.AllocationNumber = v316; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v317 => { _outer.AllocationNumber = v317; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT pda FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET pda = pda+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("PD00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("PD0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("PD000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("PD00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("PD0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("PD", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v318 => { _outer.CINumber = v318; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v319 => { _outer.NewCI = v319; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v320 => { _outer.hlCase = v320; }).Ref(_outer.NewCI, v321 => { _outer.NewCI = v321; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v322 => { _outer.QryString = v322; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent12 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent12.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent12.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v323 => { _outer.AssetGroup = v323; }).Ref(_outer.NewCI, v324 => { _outer.NewCI = v324; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v325 => { _outer.PosID = v325; }).Val((Int16)0).Val("0"));
                        //BlackBerry
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeBlackberry")))
                    {
                        var loopEnd11 = _.NUM(_outer.CIQuantity);
                        var loopStart11 = _.NUM((Int16)1, loopEnd11);
                        if (_.StrictLTE(loopStart11, loopEnd11))
                        {
                            for (_outer.i = loopStart11; _.StrictLTE(_outer.i, loopEnd11); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MobileDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v326 => { _outer.OrderNumber = v326; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v327 => { _outer.VendorName = v327; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v328 => { _outer.OrderDate = v328; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v329 => { _outer.CompanyCode = v329; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v330 => { _outer.VendorNumber = v330; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v331 => { _outer.OrderPosNr = v331; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v332 => { _outer.AllocationNumber = v332; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v333 => { _outer.AllocationType = v333; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v334 => { _outer.Reciever = v334; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v335 => { _outer.PosID = v335; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v336 => { _outer.PlaceOfUnloading = v336; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v337 => { _outer.CIComment = v337; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeBlackBerry"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v338 => { _outer.ActDate = v338; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v339 => { _outer.DeliveryDate = v339; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v340 => { _outer.ArticleDescription = v340; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v341 => { _outer.CIPrice = v341; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v342 => { _outer.CIPriceCurrency = v342; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v343 => { _outer.AllocationNumber = v343; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v344 => { _outer.AllocationNumber = v344; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT blackberry FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET blackberry = blackberry+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("BB00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("BB0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("BB000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("BB00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("BB0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("BB", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v345 => { _outer.CINumber = v345; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v346 => { _outer.NewCI = v346; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v347 => { _outer.hlCase = v347; }).Ref(_outer.NewCI, v348 => { _outer.NewCI = v348; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v349 => { _outer.QryString = v349; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent13 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent13.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent13.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v350 => { _outer.AssetGroup = v350; }).Ref(_outer.NewCI, v351 => { _outer.NewCI = v351; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v352 => { _outer.PosID = v352; }).Val((Int16)0).Val("0"));
                        //Monitor
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeMonitor")))
                    {
                        var loopEnd12 = _.NUM(_outer.CIQuantity);
                        var loopStart12 = _.NUM((Int16)1, loopEnd12);
                        if (_.StrictLTE(loopStart12, loopEnd12))
                        {
                            for (_outer.i = loopStart12; _.StrictLTE(_outer.i, loopEnd12); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "Monitor"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v353 => { _outer.OrderNumber = v353; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v354 => { _outer.VendorName = v354; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v355 => { _outer.OrderDate = v355; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v356 => { _outer.CompanyCode = v356; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v357 => { _outer.VendorNumber = v357; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v358 => { _outer.OrderPosNr = v358; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v359 => { _outer.AllocationNumber = v359; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v360 => { _outer.AllocationType = v360; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v361 => { _outer.Reciever = v361; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v362 => { _outer.PosID = v362; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v363 => { _outer.PlaceOfUnloading = v363; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v364 => { _outer.CIComment = v364; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v365 => { _outer.ActDate = v365; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v366 => { _outer.DeliveryDate = v366; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v367 => { _outer.ArticleDescription = v367; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v368 => { _outer.CIPrice = v368; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v369 => { _outer.CIPriceCurrency = v369; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v370 => { _outer.AllocationNumber = v370; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v371 => { _outer.AllocationNumber = v371; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT monitor FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET monitor = monitor+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("MO00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("MO0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("MO000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("MO00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("MO0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("MO", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v372 => { _outer.CINumber = v372; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v373 => { _outer.NewCI = v373; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v374 => { _outer.hlCase = v374; }).Ref(_outer.NewCI, v375 => { _outer.NewCI = v375; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v376 => { _outer.QryString = v376; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent14 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent14.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent14.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v377 => { _outer.AssetGroup = v377; }).Ref(_outer.NewCI, v378 => { _outer.NewCI = v378; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v379 => { _outer.PosID = v379; }).Val((Int16)0).Val("0"));
                        //Beamer
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeBeamer")))
                    {
                        var loopEnd13 = _.NUM(_outer.CIQuantity);
                        var loopStart13 = _.NUM((Int16)1, loopEnd13);
                        if (_.StrictLTE(loopStart13, loopEnd13))
                        {
                            for (_outer.i = loopStart13; _.StrictLTE(_outer.i, loopEnd13); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MultiMediaDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v380 => { _outer.OrderNumber = v380; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v381 => { _outer.VendorName = v381; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v382 => { _outer.OrderDate = v382; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v383 => { _outer.CompanyCode = v383; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v384 => { _outer.VendorNumber = v384; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v385 => { _outer.OrderPosNr = v385; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v386 => { _outer.AllocationNumber = v386; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v387 => { _outer.AllocationType = v387; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v388 => { _outer.Reciever = v388; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v389 => { _outer.PosID = v389; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v390 => { _outer.PlaceOfUnloading = v390; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v391 => { _outer.CIComment = v391; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v392 => { _outer.ActDate = v392; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v393 => { _outer.DeliveryDate = v393; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeBeamer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v394 => { _outer.ArticleDescription = v394; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v395 => { _outer.CIPrice = v395; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v396 => { _outer.CIPriceCurrency = v396; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v397 => { _outer.AllocationNumber = v397; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v398 => { _outer.AllocationNumber = v398; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT beamer FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET beamer = beamer+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("VP00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("VP0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("VP000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("VP00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("VP0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("VP", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v399 => { _outer.CINumber = v399; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v400 => { _outer.NewCI = v400; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v401 => { _outer.hlCase = v401; }).Ref(_outer.NewCI, v402 => { _outer.NewCI = v402; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v403 => { _outer.QryString = v403; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent15 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent15.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent15.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v404 => { _outer.AssetGroup = v404; }).Ref(_outer.NewCI, v405 => { _outer.NewCI = v405; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v406 => { _outer.PosID = v406; }).Val((Int16)0).Val("0"));
                        //Videokonferenztechnik
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeVideoconferencetechnic")))
                    {
                        var loopEnd14 = _.NUM(_outer.CIQuantity);
                        var loopStart14 = _.NUM((Int16)1, loopEnd14);
                        if (_.StrictLTE(loopStart14, loopEnd14))
                        {
                            for (_outer.i = loopStart14; _.StrictLTE(_outer.i, loopEnd14); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MultiMediaDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v407 => { _outer.OrderNumber = v407; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v408 => { _outer.VendorName = v408; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v409 => { _outer.OrderDate = v409; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v410 => { _outer.CompanyCode = v410; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v411 => { _outer.VendorNumber = v411; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v412 => { _outer.OrderPosNr = v412; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v413 => { _outer.AllocationNumber = v413; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v414 => { _outer.AllocationType = v414; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v415 => { _outer.Reciever = v415; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v416 => { _outer.PosID = v416; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v417 => { _outer.PlaceOfUnloading = v417; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v418 => { _outer.CIComment = v418; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v419 => { _outer.ActDate = v419; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v420 => { _outer.DeliveryDate = v420; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeVideoConferenceTechnic"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v421 => { _outer.ArticleDescription = v421; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v422 => { _outer.CIPrice = v422; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v423 => { _outer.CIPriceCurrency = v423; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v424 => { _outer.AllocationNumber = v424; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v425 => { _outer.AllocationNumber = v425; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT videoconference FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET videoconference = videoconference+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("VC00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("VC0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("VC000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("VC00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("VC0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("VC", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v426 => { _outer.CINumber = v426; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v427 => { _outer.NewCI = v427; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v428 => { _outer.hlCase = v428; }).Ref(_outer.NewCI, v429 => { _outer.NewCI = v429; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v430 => { _outer.QryString = v430; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent16 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent16.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent16.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v431 => { _outer.AssetGroup = v431; }).Ref(_outer.NewCI, v432 => { _outer.NewCI = v432; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v433 => { _outer.PosID = v433; }).Val((Int16)0).Val("0"));
                        //Medientechnik
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeMediaTechnic")))
                    {
                        var loopEnd15 = _.NUM(_outer.CIQuantity);
                        var loopStart15 = _.NUM((Int16)1, loopEnd15);
                        if (_.StrictLTE(loopStart15, loopEnd15))
                        {
                            for (_outer.i = loopStart15; _.StrictLTE(_outer.i, loopEnd15); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "MultiMediaDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v434 => { _outer.OrderNumber = v434; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v435 => { _outer.VendorName = v435; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v436 => { _outer.OrderDate = v436; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v437 => { _outer.CompanyCode = v437; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v438 => { _outer.VendorNumber = v438; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v439 => { _outer.OrderPosNr = v439; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v440 => { _outer.AllocationNumber = v440; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v441 => { _outer.AllocationType = v441; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v442 => { _outer.Reciever = v442; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v443 => { _outer.PosID = v443; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v444 => { _outer.PlaceOfUnloading = v444; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v445 => { _outer.CIComment = v445; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v446 => { _outer.ActDate = v446; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v447 => { _outer.DeliveryDate = v447; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeMediaTechnic"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v448 => { _outer.ArticleDescription = v448; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v449 => { _outer.CIPrice = v449; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v450 => { _outer.CIPriceCurrency = v450; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v451 => { _outer.AllocationNumber = v451; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v452 => { _outer.AllocationNumber = v452; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT mediatechnic FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET mediatechnic = mediatechnic+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("MU00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("MU0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("MU000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("MU00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("MU0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("MU", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v453 => { _outer.CINumber = v453; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v454 => { _outer.NewCI = v454; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v455 => { _outer.hlCase = v455; }).Ref(_outer.NewCI, v456 => { _outer.NewCI = v456; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v457 => { _outer.QryString = v457; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent17 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent17.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent17.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v458 => { _outer.AssetGroup = v458; }).Ref(_outer.NewCI, v459 => { _outer.NewCI = v459; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v460 => { _outer.PosID = v460; }).Val((Int16)0).Val("0"));
                        //Diktiergeraet
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeDictaphone")))
                    {
                        var loopEnd16 = _.NUM(_outer.CIQuantity);
                        var loopStart16 = _.NUM((Int16)1, loopEnd16);
                        if (_.StrictLTE(loopStart16, loopEnd16))
                        {
                            for (_outer.i = loopStart16; _.StrictLTE(_outer.i, loopEnd16); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v461 => { _outer.OrderNumber = v461; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v462 => { _outer.VendorName = v462; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v463 => { _outer.OrderDate = v463; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v464 => { _outer.CompanyCode = v464; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v465 => { _outer.VendorNumber = v465; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v466 => { _outer.OrderPosNr = v466; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v467 => { _outer.AllocationNumber = v467; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v468 => { _outer.AllocationType = v468; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v469 => { _outer.Reciever = v469; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v470 => { _outer.PosID = v470; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v471 => { _outer.PlaceOfUnloading = v471; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v472 => { _outer.CIComment = v472; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v473 => { _outer.ActDate = v473; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v474 => { _outer.DeliveryDate = v474; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeDictationDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v475 => { _outer.ArticleDescription = v475; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v476 => { _outer.CIPrice = v476; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v477 => { _outer.CIPriceCurrency = v477; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v478 => { _outer.AllocationNumber = v478; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v479 => { _outer.AllocationNumber = v479; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT diktiersystem FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET diktiersystem = diktiersystem+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("DS00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("DS0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("DS000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("DS00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("DS0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("DS", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v480 => { _outer.CINumber = v480; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v481 => { _outer.NewCI = v481; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v482 => { _outer.hlCase = v482; }).Ref(_outer.NewCI, v483 => { _outer.NewCI = v483; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v484 => { _outer.QryString = v484; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent18 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent18.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent18.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v485 => { _outer.AssetGroup = v485; }).Ref(_outer.NewCI, v486 => { _outer.NewCI = v486; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v487 => { _outer.PosID = v487; }).Val((Int16)0).Val("0"));
                        //USV
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeUSV")))
                    {
                        var loopEnd17 = _.NUM(_outer.CIQuantity);
                        var loopStart17 = _.NUM((Int16)1, loopEnd17);
                        if (_.StrictLTE(loopStart17, loopEnd17))
                        {
                            for (_outer.i = loopStart17; _.StrictLTE(_outer.i, loopEnd17); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v488 => { _outer.OrderNumber = v488; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v489 => { _outer.VendorName = v489; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v490 => { _outer.OrderDate = v490; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v491 => { _outer.CompanyCode = v491; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v492 => { _outer.VendorNumber = v492; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v493 => { _outer.OrderPosNr = v493; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v494 => { _outer.AllocationNumber = v494; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v495 => { _outer.AllocationType = v495; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v496 => { _outer.Reciever = v496; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v497 => { _outer.PosID = v497; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v498 => { _outer.PlaceOfUnloading = v498; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v499 => { _outer.CIComment = v499; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v500 => { _outer.ActDate = v500; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v501 => { _outer.DeliveryDate = v501; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeUSV"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v502 => { _outer.ArticleDescription = v502; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v503 => { _outer.CIPrice = v503; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v504 => { _outer.CIPriceCurrency = v504; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v505 => { _outer.AllocationNumber = v505; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v506 => { _outer.AllocationNumber = v506; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT usv FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET usv = usv+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("UP00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("UP0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("UP000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("UP00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("UP0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("UP", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v507 => { _outer.CINumber = v507; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v508 => { _outer.NewCI = v508; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v509 => { _outer.hlCase = v509; }).Ref(_outer.NewCI, v510 => { _outer.NewCI = v510; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v511 => { _outer.QryString = v511; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent19 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent19.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent19.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v512 => { _outer.AssetGroup = v512; }).Ref(_outer.NewCI, v513 => { _outer.NewCI = v513; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v514 => { _outer.PosID = v514; }).Val((Int16)0).Val("0"));
                        //ueberwachungskamera
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeControlCam")))
                    {
                        var loopEnd18 = _.NUM(_outer.CIQuantity);
                        var loopStart18 = _.NUM((Int16)1, loopEnd18);
                        if (_.StrictLTE(loopStart18, loopEnd18))
                        {
                            for (_outer.i = loopStart18; _.StrictLTE(_outer.i, loopEnd18); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v515 => { _outer.OrderNumber = v515; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v516 => { _outer.VendorName = v516; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v517 => { _outer.OrderDate = v517; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v518 => { _outer.CompanyCode = v518; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v519 => { _outer.VendorNumber = v519; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v520 => { _outer.OrderPosNr = v520; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v521 => { _outer.AllocationNumber = v521; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v522 => { _outer.AllocationType = v522; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v523 => { _outer.Reciever = v523; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v524 => { _outer.PosID = v524; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v525 => { _outer.PlaceOfUnloading = v525; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v526 => { _outer.CIComment = v526; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v527 => { _outer.ActDate = v527; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v528 => { _outer.DeliveryDate = v528; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeControlCam"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v529 => { _outer.ArticleDescription = v529; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v530 => { _outer.CIPrice = v530; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v531 => { _outer.CIPriceCurrency = v531; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v532 => { _outer.AllocationNumber = v532; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v533 => { _outer.AllocationNumber = v533; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT controlcam FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET controlcam = controlcam+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("MC00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("MC0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("MC000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("MC00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("MC0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("MC", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v534 => { _outer.CINumber = v534; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v535 => { _outer.NewCI = v535; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v536 => { _outer.hlCase = v536; }).Ref(_outer.NewCI, v537 => { _outer.NewCI = v537; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v538 => { _outer.QryString = v538; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent20 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent20.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent20.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v539 => { _outer.AssetGroup = v539; }).Ref(_outer.NewCI, v540 => { _outer.NewCI = v540; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v541 => { _outer.PosID = v541; }).Val((Int16)0).Val("0"));
                        //BDE
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeBDE")))
                    {
                        var loopEnd19 = _.NUM(_outer.CIQuantity);
                        var loopStart19 = _.NUM((Int16)1, loopEnd19);
                        if (_.StrictLTE(loopStart19, loopEnd19))
                        {
                            for (_outer.i = loopStart19; _.StrictLTE(_outer.i, loopEnd19); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v542 => { _outer.OrderNumber = v542; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v543 => { _outer.VendorName = v543; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v544 => { _outer.OrderDate = v544; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v545 => { _outer.CompanyCode = v545; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v546 => { _outer.VendorNumber = v546; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v547 => { _outer.OrderPosNr = v547; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v548 => { _outer.AllocationNumber = v548; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v549 => { _outer.AllocationType = v549; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v550 => { _outer.Reciever = v550; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v551 => { _outer.PosID = v551; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v552 => { _outer.PlaceOfUnloading = v552; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v553 => { _outer.CIComment = v553; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v554 => { _outer.ActDate = v554; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v555 => { _outer.DeliveryDate = v555; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeBDE"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v556 => { _outer.ArticleDescription = v556; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v557 => { _outer.CIPrice = v557; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v558 => { _outer.CIPriceCurrency = v558; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v559 => { _outer.AllocationNumber = v559; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v560 => { _outer.AllocationNumber = v560; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT bde FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET bde = bde+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("DA00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("DA0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("DA000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("DA00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("DA0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("DA", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v561 => { _outer.CINumber = v561; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v562 => { _outer.NewCI = v562; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v563 => { _outer.hlCase = v563; }).Ref(_outer.NewCI, v564 => { _outer.NewCI = v564; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v565 => { _outer.QryString = v565; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent21 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent21.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent21.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v566 => { _outer.AssetGroup = v566; }).Ref(_outer.NewCI, v567 => { _outer.NewCI = v567; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v568 => { _outer.PosID = v568; }).Val((Int16)0).Val("0"));
                        //Spacemaus
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeSpacemouse")))
                    {
                        var loopEnd20 = _.NUM(_outer.CIQuantity);
                        var loopStart20 = _.NUM((Int16)1, loopEnd20);
                        if (_.StrictLTE(loopStart20, loopEnd20))
                        {
                            for (_outer.i = loopStart20; _.StrictLTE(_outer.i, loopEnd20); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v569 => { _outer.OrderNumber = v569; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v570 => { _outer.VendorName = v570; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v571 => { _outer.OrderDate = v571; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v572 => { _outer.CompanyCode = v572; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v573 => { _outer.VendorNumber = v573; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v574 => { _outer.OrderPosNr = v574; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v575 => { _outer.AllocationNumber = v575; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v576 => { _outer.AllocationType = v576; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v577 => { _outer.Reciever = v577; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v578 => { _outer.PosID = v578; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v579 => { _outer.PlaceOfUnloading = v579; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v580 => { _outer.CIComment = v580; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v581 => { _outer.ActDate = v581; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v582 => { _outer.DeliveryDate = v582; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeSpaceMouse"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v583 => { _outer.ArticleDescription = v583; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v584 => { _outer.CIPrice = v584; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v585 => { _outer.CIPriceCurrency = v585; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v586 => { _outer.AllocationNumber = v586; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v587 => { _outer.AllocationNumber = v587; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT spacemouse FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET spacemouse = spacemouse+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("SP00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("SP0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("SP000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("SP00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("SP0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("SP", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v588 => { _outer.CINumber = v588; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v589 => { _outer.NewCI = v589; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v590 => { _outer.hlCase = v590; }).Ref(_outer.NewCI, v591 => { _outer.NewCI = v591; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v592 => { _outer.QryString = v592; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent22 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent22.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent22.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v593 => { _outer.AssetGroup = v593; }).Ref(_outer.NewCI, v594 => { _outer.NewCI = v594; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v595 => { _outer.PosID = v595; }).Val((Int16)0).Val("0"));
                        //Aktive Netzwerkkomponente
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeNetworkComponent")))
                    {
                        var loopEnd21 = _.NUM(_outer.CIQuantity);
                        var loopStart21 = _.NUM((Int16)1, loopEnd21);
                        if (_.StrictLTE(loopStart21, loopEnd21))
                        {
                            for (_outer.i = loopStart21; _.StrictLTE(_outer.i, loopEnd21); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "NetworkComponent"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v596 => { _outer.OrderNumber = v596; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v597 => { _outer.VendorName = v597; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v598 => { _outer.OrderDate = v598; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v599 => { _outer.CompanyCode = v599; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v600 => { _outer.VendorNumber = v600; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v601 => { _outer.OrderPosNr = v601; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v602 => { _outer.AllocationNumber = v602; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v603 => { _outer.AllocationType = v603; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v604 => { _outer.Reciever = v604; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v605 => { _outer.PosID = v605; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v606 => { _outer.PlaceOfUnloading = v606; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v607 => { _outer.CIComment = v607; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v608 => { _outer.ActDate = v608; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v609 => { _outer.DeliveryDate = v609; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("NetworkComponentDetail.NetworkComponentType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("TypeActiveNetworkComponet"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v610 => { _outer.ArticleDescription = v610; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v611 => { _outer.CIPrice = v611; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v612 => { _outer.CIPriceCurrency = v612; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v613 => { _outer.AllocationNumber = v613; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v614 => { _outer.AllocationNumber = v614; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT networkcomponent FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET networkcomponent = networkcomponent+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("AN00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("AN0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("AN000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("AN00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("AN0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("AN", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v615 => { _outer.CINumber = v615; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v616 => { _outer.NewCI = v616; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v617 => { _outer.hlCase = v617; }).Ref(_outer.NewCI, v618 => { _outer.NewCI = v618; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v619 => { _outer.QryString = v619; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent23 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent23.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent23.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v620 => { _outer.AssetGroup = v620; }).Ref(_outer.NewCI, v621 => { _outer.NewCI = v621; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v622 => { _outer.PosID = v622; }).Val((Int16)0).Val("0"));
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeHomeOfficeRouter")))
                    {
                        var loopEnd22 = _.NUM(_outer.CIQuantity);
                        var loopStart22 = _.NUM((Int16)1, loopEnd22);
                        if (_.StrictLTE(loopStart22, loopEnd22))
                        {
                            for (_outer.i = loopStart22; _.StrictLTE(_outer.i, loopEnd22); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "NetworkComponent"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v623 => { _outer.OrderNumber = v623; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v624 => { _outer.VendorName = v624; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v625 => { _outer.OrderDate = v625; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v626 => { _outer.CompanyCode = v626; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v627 => { _outer.VendorNumber = v627; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v628 => { _outer.OrderPosNr = v628; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v629 => { _outer.AllocationNumber = v629; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v630 => { _outer.AllocationType = v630; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v631 => { _outer.Reciever = v631; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v632 => { _outer.PosID = v632; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v633 => { _outer.PlaceOfUnloading = v633; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v634 => { _outer.CIComment = v634; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v635 => { _outer.ActDate = v635; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v636 => { _outer.DeliveryDate = v636; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("NetworkComponentDetail.NetworkComponentType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("TypeHomeOfficeRouter"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v637 => { _outer.ArticleDescription = v637; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v638 => { _outer.CIPrice = v638; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v639 => { _outer.CIPriceCurrency = v639; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v640 => { _outer.AllocationNumber = v640; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v641 => { _outer.AllocationNumber = v641; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT homeofficerouter FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET homeofficerouter = homeofficerouter+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("HO00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("HO0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("HO000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("HO00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("HO0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("HO", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v642 => { _outer.CINumber = v642; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v643 => { _outer.NewCI = v643; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v644 => { _outer.hlCase = v644; }).Ref(_outer.NewCI, v645 => { _outer.NewCI = v645; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v646 => { _outer.QryString = v646; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent24 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent24.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent24.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v647 => { _outer.AssetGroup = v647; }).Ref(_outer.NewCI, v648 => { _outer.NewCI = v648; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v649 => { _outer.PosID = v649; }).Val((Int16)0).Val("0"));
                        //Headset
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeHeadset")))
                    {
                        var loopEnd23 = _.NUM(_outer.CIQuantity);
                        var loopStart23 = _.NUM((Int16)1, loopEnd23);
                        if (_.StrictLTE(loopStart23, loopEnd23))
                        {
                            for (_outer.i = loopStart23; _.StrictLTE(_outer.i, loopEnd23); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v650 => { _outer.OrderNumber = v650; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v651 => { _outer.VendorName = v651; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v652 => { _outer.OrderDate = v652; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v653 => { _outer.CompanyCode = v653; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v654 => { _outer.VendorNumber = v654; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v655 => { _outer.OrderPosNr = v655; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v656 => { _outer.AllocationNumber = v656; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v657 => { _outer.AllocationType = v657; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v658 => { _outer.Reciever = v658; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v659 => { _outer.PosID = v659; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v660 => { _outer.PlaceOfUnloading = v660; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v661 => { _outer.CIComment = v661; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v662 => { _outer.ActDate = v662; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v663 => { _outer.DeliveryDate = v663; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeHeadset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v664 => { _outer.ArticleDescription = v664; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v665 => { _outer.CIPrice = v665; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v666 => { _outer.CIPriceCurrency = v666; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v667 => { _outer.AllocationNumber = v667; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v668 => { _outer.AllocationNumber = v668; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT headset FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET headset = headset+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("HS00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("HS0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("HS000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("HS00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("HS0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("HS", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v669 => { _outer.CINumber = v669; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v670 => { _outer.NewCI = v670; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v671 => { _outer.hlCase = v671; }).Ref(_outer.NewCI, v672 => { _outer.NewCI = v672; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v673 => { _outer.QryString = v673; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent25 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent25.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent25.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v674 => { _outer.AssetGroup = v674; }).Ref(_outer.NewCI, v675 => { _outer.NewCI = v675; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v676 => { _outer.PosID = v676; }).Val((Int16)0).Val("0"));

                        //ConferencePhone
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeConferencePhone")))
                    {
                        var loopEnd24 = _.NUM(_outer.CIQuantity);
                        var loopStart24 = _.NUM((Int16)1, loopEnd24);
                        if (_.StrictLTE(loopStart24, loopEnd24))
                        {
                            for (_outer.i = loopStart24; _.StrictLTE(_outer.i, loopEnd24); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Notebook anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "GenericAsset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v677 => { _outer.OrderNumber = v677; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v678 => { _outer.VendorName = v678; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v679 => { _outer.OrderDate = v679; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v680 => { _outer.CompanyCode = v680; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v681 => { _outer.VendorNumber = v681; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v682 => { _outer.OrderPosNr = v682; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v683 => { _outer.AllocationNumber = v683; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v684 => { _outer.AllocationType = v684; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v685 => { _outer.Reciever = v685; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v686 => { _outer.PosID = v686; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v687 => { _outer.PlaceOfUnloading = v687; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v688 => { _outer.CIComment = v688; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v689 => { _outer.ActDate = v689; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v690 => { _outer.DeliveryDate = v690; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeConferencePhone"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v691 => { _outer.ArticleDescription = v691; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v692 => { _outer.CIPrice = v692; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v693 => { _outer.CIPriceCurrency = v693; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v694 => { _outer.AllocationNumber = v694; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v695 => { _outer.AllocationNumber = v695; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm1", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT conferencephone FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET conferencephone = conferencephone+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("CP00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("CP0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("CP000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("CP00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("CP0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("CP", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v696 => { _outer.CINumber = v696; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v697 => { _outer.NewCI = v697; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v698 => { _outer.hlCase = v698; }).Ref(_outer.NewCI, v699 => { _outer.NewCI = v699; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v700 => { _outer.QryString = v700; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent26 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent26.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent26.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v701 => { _outer.AssetGroup = v701; }).Ref(_outer.NewCI, v702 => { _outer.NewCI = v702; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v703 => { _outer.PosID = v703; }).Val((Int16)0).Val("0"));

                        //Server
                    }
                    else if (_.IF(_.EQ(_outer.CIType, "CITypeServerComputer")))
                    {
                        var loopEnd25 = _.NUM(_outer.CIQuantity);
                        var loopStart25 = _.NUM((Int16)1, loopEnd25);
                        if (_.StrictLTE(loopStart25, loopEnd25))
                        {
                            for (_outer.i = loopStart25; _.StrictLTE(_outer.i, loopEnd25); _outer.i = _.ADD(_outer.i, (Int16)1))
                            {
                                //hlContext.Trace 1, "Computer anlegen Nummer: " & i
                                _outer.NewCI = _.OBJ(_.CALLm1v1(this, _env.hlContext, "createobject", "ServerComputer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v704 => { _outer.OrderNumber = v704; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v705 => { _outer.VendorName = v705; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v706 => { _outer.OrderDate = v706; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v707 => { _outer.CompanyCode = v707; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v708 => { _outer.VendorNumber = v708; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v709 => { _outer.OrderPosNr = v709; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v710 => { _outer.AllocationNumber = v710; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v711 => { _outer.AllocationType = v711; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v712 => { _outer.Reciever = v712; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v713 => { _outer.PosID = v713; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v714 => { _outer.PlaceOfUnloading = v714; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v715 => { _outer.CIComment = v715; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v716 => { _outer.ActDate = v716; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v717 => { _outer.DeliveryDate = v717; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v718 => { _outer.ArticleDescription = v718; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v719 => { _outer.CIPrice = v719; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v720 => { _outer.CIPriceCurrency = v720; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v721 => { _outer.AllocationNumber = v721; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v722 => { _outer.AllocationNumber = v722; }));
                                }
                                //------------------------------------------------------------------------------------------------
                                //Generiert eine neue CI-Nummer
                                _outer.cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                                //Verbindung oeffnen
                                //Hier muss Server- und Datenbankname fest eingetragen werden!
                                //Wird die DB auf einen anderen Server uebertragen, muss dies vor Betrieb hier angepasst werden!!!
                                _.SET("Provider=SQLOLEDB.1;Password=helplineuser;Persist Security Info=True;User ID=helplineuser;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, _outer.cn, "ConnectionString");
                                _.SET((Int16)10, this, _outer.cn, "ConnectionTimeout");
                                _.CALLm1v0(this, _outer.cn, "Open");

                                //CI-Nummer auslesen
                                _outer.rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                                _outer.rs = _.OBJ(_.CALLm1v1(this, _outer.cn, "Execute", "SELECT server FROM _cinumbers"));
                                //In Variable schreiben
                                _outer.CINumber = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, _outer.rs, "fields", (Int16)0), "value"));
                                //CI-Nummer in der Datenbank um den Wert 1 erhoehen und zurueckschreiben
                                _.CALLm1v1(this, _outer.cn, "execute", "UPDATE _cinumbers SET server = server+1");

                                //Verbindung schliessen
                                _.CALLm1v0(this, _outer.rs, "close");
                                _.CALLm1v0(this, _outer.cn, "close");
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)1)))
                                {
                                    _outer.CINumber = _.ADD("SR00000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)2)))
                                {
                                    _outer.CINumber = _.ADD("SR0000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)3)))
                                {
                                    _outer.CINumber = _.ADD("SR000", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)4)))
                                {
                                    _outer.CINumber = _.ADD("SR00", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)5)))
                                {
                                    _outer.CINumber = _.ADD("SR0", _.CSTR(_outer.CINumber));
                                }
                                if (_.IF(_.EQ(_.NullableNUM(_.LEN(_.CSTR(_outer.CINumber))), (Int16)6)))
                                {
                                    _outer.CINumber = _.ADD("SR", _.CSTR(_outer.CINumber));
                                }

                                //hlContext.Trace 1, "Ticket-ID = " & NextTicketID
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v723 => { _outer.CINumber = v723; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v724 => { _outer.NewCI = v724; }));
                                //Neues CI dem Vorgang assoziieren
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v725 => { _outer.hlCase = v725; }).Ref(_outer.NewCI, v726 => { _outer.NewCI = v726; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                _.CALLm1v2(this, _env.hlContext, "Trace", (Int16)1, _.CONCAT("Suche Inv-Gruppe: ", _outer.QryString));
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v727 => { _outer.QryString = v727; })));
                                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v2(this, _outer.Qry, "GetItemCount", (Int16)0, (Int16)0)), "1")))
                                {
                                    _outer.AssetGroups = _.VAL(_.CALLm1argp(this, _outer.Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
                                    var enumerationContent27 = _.ENUMERABLE(_outer.AssetGroups).GetEnumerator();
                                    while (true)
                                    {
                                        if (!enumerationContent27.MoveNext())
                                            break;
                                        _outer.rewritten_Group = enumerationContent27.Current;
                                        _outer.AssetGroup = _.OBJ(_outer.rewritten_Group);
                                    }
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v728 => { _outer.AssetGroup = v728; }).Ref(_outer.NewCI, v729 => { _outer.NewCI = v729; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v730 => { _outer.PosID = v730; }).Val((Int16)0).Val("0"));
                    }
                    //Kennzeichnen, dass CI erzeugt wurde
                    _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIisCreated").Val((Int16)0).Ref(_outer.PosID, v731 => { _outer.PosID = v731; }).Val((Int16)0).Val("1"));
                }
                _outer.CreateCI = (Int16)0;
            }

            //Eliminierung von Geraeten------------------------------------------------------------------------------------------------------
            //Pruefen, ob Anzahl der assoziierten CI's pro Typ groesser ist, als Inhalt des Attributs Bestellmenge
            //Wenn ja, dann entsprechend viele CI's (Differenz aus Anzahl und Bestellmenge) auf Status "eliminiert" setzen
            //Das Ganze nur bei OrderStatus Aenderung/Storno
            _outer.objs = VBScriptConstants.Nothing;
            _outer.obj = VBScriptConstants.Nothing;
            _outer.objtype = VBScriptConstants.Nothing;
            _outer.cistatus = "";
            _outer.statuscounter = (Int16)0;
            _outer.typecounter = (Int16)0;
            _outer.stornoquantity = (Int16)0;
            _outer.stornocounter = (Int16)0;
            _outer.statusoverview = "";
            _outer.CIExistingAtSAPAM = (Int16)0;
            _outer.OrderPosID = (Int16)0;
            _outer.PosType = "";
            var enumerationContent28 = _.ENUMERABLE(_outer.OrderPosIDs).GetEnumerator();
            while (true)
            {
                if (!enumerationContent28.MoveNext())
                    break;
                _outer.PosID = enumerationContent28.Current;
                //Pruefen ob CI erzeugt werden soll
                _outer.CreateCI = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CreateCI").Val((Int16)0).Ref(_outer.PosID, v732 => { _outer.PosID = v732; }).Val((Int16)0).Val((Int16)0)));
                //Pruefen ob CI bereits erzeugt wurde
                _outer.CIisCreated = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIisCreated").Val((Int16)0).Ref(_outer.PosID, v733 => { _outer.PosID = v733; }).Val((Int16)0).Val((Int16)0)));
                //Geraetetyp validieren
                _outer.PosType = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIType").Val((Int16)0).Ref(_outer.PosID, v734 => { _outer.PosID = v734; }).Val((Int16)0).Val((Int16)0)));
                //Bestellmenge abfragen
                _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v735 => { _outer.PosID = v735; }).Val((Int16)0).Val((Int16)0)));
                //Auf Eliminierung pruefen
                if (_.IF(_.AND(_.AND(_.EQ(_.NullableSTR(_outer.CreateCI), "1"), _.EQ(_.NullableSTR(_outer.CIisCreated), "1")), _.EQ(_.NullableSTR(_outer.strOrdReqStatus), "OrderRequestStatusChangedStorno"))))
                {
                    //Anzahl assoziierte CIs ermitteln
                    _outer.objs = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val(119155)));
                    //DesktopComputer
                    var enumerationContent29 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent29.MoveNext())
                            break;
                        _outer.obj = enumerationContent29.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "DesktopComputer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeDesktopcomputer")))
                        {
                            var enumerationContent30 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent30.MoveNext())
                                    break;
                                _outer.obj = enumerationContent30.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "DesktopComputer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v736 => { _outer.ActDate = v736; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v737 => { _outer.statusoverview = v737; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v738 => { _env.hlContext = v738; }).Ref(_outer.obj, v739 => { _outer.obj = v739; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v740 => { _outer.obj = v740; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //NotebookComputer
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent31 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent31.MoveNext())
                            break;
                        _outer.obj = enumerationContent31.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "NotebookComputer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeNotebook")))
                        {
                            var enumerationContent32 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent32.MoveNext())
                                    break;
                                _outer.obj = enumerationContent32.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "NotebookComputer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v741 => { _outer.ActDate = v741; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v742 => { _outer.statusoverview = v742; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v743 => { _env.hlContext = v743; }).Ref(_outer.obj, v744 => { _outer.obj = v744; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v745 => { _outer.obj = v745; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Aktive Netzwerkomponente
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent33 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent33.MoveNext())
                            break;
                        _outer.obj = enumerationContent33.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "NetworkComponent")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeNetworkComponent")))
                        {
                            var enumerationContent34 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent34.MoveNext())
                                    break;
                                _outer.obj = enumerationContent34.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "NotebookComputer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v746 => { _outer.ActDate = v746; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v747 => { _outer.statusoverview = v747; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v748 => { _env.hlContext = v748; }).Ref(_outer.obj, v749 => { _outer.obj = v749; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v750 => { _outer.obj = v750; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Monitor
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent35 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent35.MoveNext())
                            break;
                        _outer.obj = enumerationContent35.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Monitor")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v751 => { _outer.PosID = v751; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeMonitor")))
                        {
                            var enumerationContent36 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent36.MoveNext())
                                    break;
                                _outer.obj = enumerationContent36.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Monitor")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v752 => { _outer.ActDate = v752; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v753 => { _outer.statusoverview = v753; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v754 => { _env.hlContext = v754; }).Ref(_outer.obj, v755 => { _outer.obj = v755; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v756 => { _outer.obj = v756; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Printer
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent37 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent37.MoveNext())
                            break;
                        _outer.obj = enumerationContent37.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v757 => { _outer.PosID = v757; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypePrinter")))
                        {
                            var enumerationContent38 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent38.MoveNext())
                                    break;
                                _outer.obj = enumerationContent38.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v758 => { _outer.ActDate = v758; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v759 => { _outer.statusoverview = v759; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v760 => { _env.hlContext = v760; }).Ref(_outer.obj, v761 => { _outer.obj = v761; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v762 => { _outer.obj = v762; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Scanner
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent39 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent39.MoveNext())
                            break;
                        _outer.obj = enumerationContent39.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v763 => { _outer.PosID = v763; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeScanner")))
                        {
                            var enumerationContent40 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent40.MoveNext())
                                    break;
                                _outer.obj = enumerationContent40.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v764 => { _outer.ActDate = v764; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v765 => { _outer.statusoverview = v765; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v766 => { _env.hlContext = v766; }).Ref(_outer.obj, v767 => { _outer.obj = v767; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v768 => { _outer.obj = v768; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Kopierer
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent41 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent41.MoveNext())
                            break;
                        _outer.obj = enumerationContent41.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v769 => { _outer.PosID = v769; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeCopyDevice")))
                        {
                            var enumerationContent42 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent42.MoveNext())
                                    break;
                                _outer.obj = enumerationContent42.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v770 => { _outer.ActDate = v770; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v771 => { _outer.statusoverview = v771; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v772 => { _env.hlContext = v772; }).Ref(_outer.obj, v773 => { _outer.obj = v773; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v774 => { _outer.obj = v774; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Multifunktionsgeraet
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent43 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent43.MoveNext())
                            break;
                        _outer.obj = enumerationContent43.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v775 => { _outer.PosID = v775; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeMultifunctionDevice")))
                        {
                            var enumerationContent44 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent44.MoveNext())
                                    break;
                                _outer.obj = enumerationContent44.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "Printer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v776 => { _outer.ActDate = v776; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v777 => { _outer.statusoverview = v777; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v778 => { _env.hlContext = v778; }).Ref(_outer.obj, v779 => { _outer.obj = v779; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v780 => { _outer.obj = v780; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Diktiergeraet
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent45 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent45.MoveNext())
                            break;
                        _outer.obj = enumerationContent45.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v781 => { _outer.PosID = v781; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeDictaphone")))
                        {
                            var enumerationContent46 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent46.MoveNext())
                                    break;
                                _outer.obj = enumerationContent46.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v782 => { _outer.ActDate = v782; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v783 => { _outer.statusoverview = v783; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v784 => { _env.hlContext = v784; }).Ref(_outer.obj, v785 => { _outer.obj = v785; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v786 => { _outer.obj = v786; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Headset
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent47 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent47.MoveNext())
                            break;
                        _outer.obj = enumerationContent47.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v787 => { _outer.PosID = v787; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeHeadset")))
                        {
                            var enumerationContent48 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent48.MoveNext())
                                    break;
                                _outer.obj = enumerationContent48.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v788 => { _outer.ActDate = v788; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v789 => { _outer.statusoverview = v789; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v790 => { _env.hlContext = v790; }).Ref(_outer.obj, v791 => { _outer.obj = v791; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v792 => { _outer.obj = v792; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //ConferencePhone
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent49 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent49.MoveNext())
                            break;
                        _outer.obj = enumerationContent49.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v793 => { _outer.PosID = v793; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeConferencePhone")))
                        {
                            var enumerationContent50 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent50.MoveNext())
                                    break;
                                _outer.obj = enumerationContent50.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v794 => { _outer.ActDate = v794; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v795 => { _outer.statusoverview = v795; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v796 => { _env.hlContext = v796; }).Ref(_outer.obj, v797 => { _outer.obj = v797; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v798 => { _outer.obj = v798; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Spacemaus
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent51 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent51.MoveNext())
                            break;
                        _outer.obj = enumerationContent51.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v799 => { _outer.PosID = v799; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeSpacemouse")))
                        {
                            var enumerationContent52 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent52.MoveNext())
                                    break;
                                _outer.obj = enumerationContent52.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v800 => { _outer.ActDate = v800; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v801 => { _outer.statusoverview = v801; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v802 => { _env.hlContext = v802; }).Ref(_outer.obj, v803 => { _outer.obj = v803; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v804 => { _outer.obj = v804; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //USV
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent53 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent53.MoveNext())
                            break;
                        _outer.obj = enumerationContent53.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v805 => { _outer.PosID = v805; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeUSV")))
                        {
                            var enumerationContent54 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent54.MoveNext())
                                    break;
                                _outer.obj = enumerationContent54.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v806 => { _outer.ActDate = v806; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v807 => { _outer.statusoverview = v807; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v808 => { _env.hlContext = v808; }).Ref(_outer.obj, v809 => { _outer.obj = v809; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v810 => { _outer.obj = v810; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Ueberwachungskamera
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent55 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent55.MoveNext())
                            break;
                        _outer.obj = enumerationContent55.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v811 => { _outer.PosID = v811; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeControlCam")))
                        {
                            var enumerationContent56 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent56.MoveNext())
                                    break;
                                _outer.obj = enumerationContent56.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v812 => { _outer.ActDate = v812; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v813 => { _outer.statusoverview = v813; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v814 => { _env.hlContext = v814; }).Ref(_outer.obj, v815 => { _outer.obj = v815; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v816 => { _outer.obj = v816; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //BDE
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent57 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent57.MoveNext())
                            break;
                        _outer.obj = enumerationContent57.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v817 => { _outer.PosID = v817; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeBDE")))
                        {
                            var enumerationContent58 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent58.MoveNext())
                                    break;
                                _outer.obj = enumerationContent58.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "GenericAsset")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v818 => { _outer.ActDate = v818; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v819 => { _outer.statusoverview = v819; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v820 => { _env.hlContext = v820; }).Ref(_outer.obj, v821 => { _outer.obj = v821; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v822 => { _outer.obj = v822; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Handy
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent59 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent59.MoveNext())
                            break;
                        _outer.obj = enumerationContent59.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v823 => { _outer.PosID = v823; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeMobilePhone")))
                        {
                            var enumerationContent60 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent60.MoveNext())
                                    break;
                                _outer.obj = enumerationContent60.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v824 => { _outer.ActDate = v824; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v825 => { _outer.statusoverview = v825; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v826 => { _env.hlContext = v826; }).Ref(_outer.obj, v827 => { _outer.obj = v827; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v828 => { _outer.obj = v828; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //BlackBerry
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent61 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent61.MoveNext())
                            break;
                        _outer.obj = enumerationContent61.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v829 => { _outer.PosID = v829; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeBlackberry")))
                        {
                            var enumerationContent62 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent62.MoveNext())
                                    break;
                                _outer.obj = enumerationContent62.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v830 => { _outer.ActDate = v830; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v831 => { _outer.statusoverview = v831; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v832 => { _env.hlContext = v832; }).Ref(_outer.obj, v833 => { _outer.obj = v833; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v834 => { _outer.obj = v834; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //PDA
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent63 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent63.MoveNext())
                            break;
                        _outer.obj = enumerationContent63.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v835 => { _outer.PosID = v835; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypePDA")))
                        {
                            var enumerationContent64 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent64.MoveNext())
                                    break;
                                _outer.obj = enumerationContent64.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v836 => { _outer.ActDate = v836; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v837 => { _outer.statusoverview = v837; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v838 => { _env.hlContext = v838; }).Ref(_outer.obj, v839 => { _outer.obj = v839; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v840 => { _outer.obj = v840; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //SIM-Karte
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent65 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent65.MoveNext())
                            break;
                        _outer.obj = enumerationContent65.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v841 => { _outer.PosID = v841; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeSIMCard")))
                        {
                            var enumerationContent66 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent66.MoveNext())
                                    break;
                                _outer.obj = enumerationContent66.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v842 => { _outer.ActDate = v842; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v843 => { _outer.statusoverview = v843; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v844 => { _env.hlContext = v844; }).Ref(_outer.obj, v845 => { _outer.obj = v845; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v846 => { _outer.obj = v846; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //UMTS-Karte
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent67 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent67.MoveNext())
                            break;
                        _outer.obj = enumerationContent67.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v847 => { _outer.PosID = v847; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeUMTSCard")))
                        {
                            var enumerationContent68 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent68.MoveNext())
                                    break;
                                _outer.obj = enumerationContent68.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MobileDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v848 => { _outer.ActDate = v848; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v849 => { _outer.statusoverview = v849; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v850 => { _env.hlContext = v850; }).Ref(_outer.obj, v851 => { _outer.obj = v851; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v852 => { _outer.obj = v852; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Videokonferenztechnik
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent69 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent69.MoveNext())
                            break;
                        _outer.obj = enumerationContent69.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v853 => { _outer.PosID = v853; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeVideoconferencetechnic")))
                        {
                            var enumerationContent70 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent70.MoveNext())
                                    break;
                                _outer.obj = enumerationContent70.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v854 => { _outer.ActDate = v854; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v855 => { _outer.statusoverview = v855; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v856 => { _env.hlContext = v856; }).Ref(_outer.obj, v857 => { _outer.obj = v857; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v858 => { _outer.obj = v858; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Beamer
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent71 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent71.MoveNext())
                            break;
                        _outer.obj = enumerationContent71.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v859 => { _outer.PosID = v859; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeBeamer")))
                        {
                            var enumerationContent72 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent72.MoveNext())
                                    break;
                                _outer.obj = enumerationContent72.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v860 => { _outer.ActDate = v860; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v861 => { _outer.statusoverview = v861; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v862 => { _env.hlContext = v862; }).Ref(_outer.obj, v863 => { _outer.obj = v863; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v864 => { _outer.obj = v864; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //Medientechnik
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent73 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent73.MoveNext())
                            break;
                        _outer.obj = enumerationContent73.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v865 => { _outer.PosID = v865; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeMediaTechnic")))
                        {
                            var enumerationContent74 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent74.MoveNext())
                                    break;
                                _outer.obj = enumerationContent74.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "MultiMediaDevice")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v866 => { _outer.ActDate = v866; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v867 => { _outer.statusoverview = v867; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v868 => { _env.hlContext = v868; }).Ref(_outer.obj, v869 => { _outer.obj = v869; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v870 => { _outer.obj = v870; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //ServerComputer
                    _outer.statuscounter = (Int16)0;
                    _outer.typecounter = (Int16)0;
                    _outer.stornocounter = (Int16)0;
                    var enumerationContent75 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent75.MoveNext())
                            break;
                        _outer.obj = enumerationContent75.Current;
                        _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "ServerComputer")))
                        {
                            _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                            {
                                //Ist Geraet eliminiert?
                                _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.cistatus), "1")))
                                {
                                    _outer.statuscounter = _.ADD(_outer.statuscounter, (Int16)1);
                                }
                                _outer.typecounter = _.ADD(_outer.typecounter, (Int16)1);
                            }
                        }
                    }
                    //Anzahl der zu eliminierenden CI`s ermitteln
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v871 => { _outer.PosID = v871; }).Val((Int16)0).Val((Int16)0)));
                    _outer.stornoquantity = _.SUBT(_.SUBT(_outer.typecounter, _outer.CIQuantity), _outer.statuscounter);
                    if (_.IF(_.GTE(_.NullableSTR(_outer.stornoquantity), "1")))
                    {
                        //Jetzt die CIs eliminieren
                        if (_.IF(_.EQ(_.NullableSTR(_outer.PosType), "CITypeServerComputer")))
                        {
                            var enumerationContent76 = _.ENUMERABLE(_outer.objs).GetEnumerator();
                            while (true)
                            {
                                if (!enumerationContent76.MoveNext())
                                    break;
                                _outer.obj = enumerationContent76.Current;
                                _outer.objtype = _.VAL(_.CALLm1argp(this, _outer.obj, "GetType", _.ARGS.ForceBrackets()));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.objtype), "ServerComputer")))
                                {
                                    //OrderPosID des Geraets abfragen
                                    _outer.OrderPosID = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                    if (_.IF(_.EQ(_.CLNG(_outer.OrderPosID), _.CLNG(_outer.PosID))))
                                    {
                                        //Ist Geraet eliminiert?
                                        _outer.cistatus = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                        if (_.IF(_.NOTEQ(_.NullableSTR(_outer.cistatus), "1")))
                                        {
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("CISubStatusStorno"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v872 => { _outer.ActDate = v872; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v873 => { _outer.statusoverview = v873; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v874 => { _env.hlContext = v874; }).Ref(_outer.obj, v875 => { _outer.obj = v875; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v876 => { _outer.obj = v876; }));
                                            _outer.stornocounter = _.ADD(_outer.stornocounter, (Int16)1);
                                        }
                                    }
                                    if (_.IF(_.EQ(_.CLNG(_outer.stornocounter), _.CLNG(_outer.stornoquantity))))
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //----------------------------------------------------------------------------------------------------------
            //Beschreibung in das SU-Attribut CaseDescriptionSU kopieren.
            //Copy Description to SU-Attribute CaseDescriptionSU.
            //Die Indizes der SUs werden festgestellt
            _outer.suindices = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetSvcUnitIndices", _.ARGS.ForceBrackets()));
            _outer.sumin = _.ADD(_.LBOUND(_outer.suindices), (Int16)1);
            _outer.sumindescr = "";
            _outer.Agent = "";
            _outer.Agent1 = "";
            _outer.Last1SUIdx = (Int16)0;
            _outer.LastSU = "";
            //Index letzte SU
            _outer.Last1SUIdx = _.VAL(_.CALLm1argp(this, _env.hlITIL2, "GetLastSUIdx", _.ARGS.Ref(_outer.hlCase, v877 => { _outer.hlCase = v877; }).Ref(_env.hlContext, v878 => { _env.hlContext = v878; })));
            //Index vorletzte SU
            _outer.LastSU = _.SUBT(_outer.Last1SUIdx, (Int16)1);
            _outer.DescrText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            //Urspruenlichen Beschreibungstext ermitteln
            _outer.sumindescr = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestSUAttribute.CaseDescriptionSU").Val((Int16)0).Val((Int16)0).Ref(_outer.sumin, v879 => { _outer.sumin = v879; }).Val((Int16)0)));
            _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v880 => { _outer.Last1SUIdx = v880; }).Val((Int16)0)));
            _outer.Agent1 = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.sumin, v881 => { _outer.sumin = v881; }).Val((Int16)0)));

            //----------------------------------------------------------------------------------------------------------
            //Kumuliert die Texte der Bearbeitungsschritte und schreibt sie in das
            //Overview-Textfeld. Die Texte werden durch Trennzeichen voneinander abgegrenzt.
            //Pruefen ob mehr als 1 SU
            _outer.DescrTextalt = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestSUAttribute.CaseDescriptionSU").Val((Int16)0).Val((Int16)0).Ref(_outer.LastSU, v882 => { _outer.LastSU = v882; }).Val((Int16)0)));
            if (_.IF(_.GT(_.NullableNUM(_outer.LastSU), (Int16)0)))
            {
                //Pruefen, ob Beschreibungstext sich geaendert hat
                if (_.IF(_.NOTEQ(_outer.DescrText, _outer.sumindescr)))
                {
                    _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v883 => { _outer.Last1SUIdx = v883; }).Val((Int16)0)));
                    _outer.DescriptionAll = "";
                    _outer.ProblemAll = "";
                    _outer.ProblemAll1 = "";
                    _outer.DiagnosisAll = "";
                    _outer.SolutionAll = "";
                    if (_.IF(_.EQ(_.NullableNUM(_outer.LangID), (Int16)7)))
                    {
                        _outer.ProblemtitleNew = _.CONCAT("=== Bestellbeschreibung neu ===", " [von Agent : ", _outer.Agent, "]", VBScriptConstants.vbNewLine);
                        _outer.Problemtitle = _.CONCAT("=== Urspruengliche Bestellbeschreibung ===", " [von Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine);
                        _outer.Diagnosistitle = _.CONCAT("=== Taetigkeitsbeschreibungen ===", VBScriptConstants.vbNewLine);
                        _outer.Solutiontitle = _.CONCAT("=== Loesungsbeschreibung ===", " [von Agent : ", _outer.Agent, "]", VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine);
                    }
                    else
                    {
                        _outer.ProblemtitleNew = _.CONCAT("=== Orderdescription new===", " [by Agent : ", _outer.Agent, "]", VBScriptConstants.vbNewLine);
                        _outer.Problemtitle = _.CONCAT("=== Original Orderdescription ===", " [by Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine);
                        _outer.Diagnosistitle = _.CONCAT("=== Diagnosisactivities ===", VBScriptConstants.vbNewLine);
                        _outer.Solutiontitle = _.CONCAT("=== Final solution ===", " [by Agent : ", _outer.Agent, "]", VBScriptConstants.vbNewLine);
                    }
                    //Problem-, Diagnose- und Loesungstext auslesen und zusammenfassen
                    _outer.Problem = _.VAL(_outer.DescrText);
                    _outer.Problem = _.REPLACE(_outer.Problem, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " ");
                    if (_.IF(_.NOTEQ(_.NullableSTR(_outer.Problem), "")))
                    {
                        _outer.ProblemAll1 = _.CONCAT(_outer.ProblemtitleNew, _outer.Problem, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                    }
                    if (_.IF(_.NOTEQ(_.NullableSTR(_outer.sumindescr), "")))
                    {
                        _outer.ProblemAll = _.CONCAT(_outer.Problemtitle, _outer.sumindescr, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                    }

                    var enumerationContent77 = _.ENUMERABLE(_outer.suindices).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent77.MoveNext())
                            break;
                        _outer.SUIdx = enumerationContent77.Current;
                        _outer.SUDiagnosis = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v884 => { _outer.SUIdx = v884; }).Val((Int16)0)));
                        //SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
                        _outer.SURegTime = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v885 => { _outer.SUIdx = v885; }).Val((Int16)0)));
                        _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v886 => { _outer.SUIdx = v886; }).Val((Int16)0)));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.SUDiagnosis), "")))
                        {
                            if (_.IF(_.EQ(_.NullableNUM(_outer.LangID), (Int16)7)))
                            {
                                _outer.SUDiagnosis = "<keine Beschreibung>";
                            }
                            else
                            {
                                _outer.SUDiagnosis = "<no description>";
                            }
                        }
                        _outer.DiagnosisAll = _.CONCAT(_outer.DiagnosisAll, _outer.SUIdx, ". SU (", _outer.Agent, ") -> [", _outer.SURegTime, "]:", VBScriptConstants.vbNewLine, _outer.SUDiagnosis, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                    }
                    _outer.Solution = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    _outer.Solution = _.REPLACE(_outer.Solution, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " ");
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.LTRIM(_.RTRIM(_outer.Solution))), "")))
                    {
                        _outer.SolutionAll = _.CONCAT(_outer.SolutionAll, _outer.Solutiontitle, _outer.Solution);
                    }
                    //Gesammelte Texte in das uebersicht-Textfeld schreiben
                    _outer.DescriptionAll = _.CONCAT(_outer.ProblemAll, _outer.ProblemAll1, _outer.Diagnosistitle, _outer.DiagnosisAll, _outer.SolutionAll);
                    _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DescriptionAll, v887 => { _outer.DescriptionAll = v887; }));
                }
            }
            if (_.IF(_.EQ(_outer.sumindescr, _outer.DescrText)))
            {
                _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v888 => { _outer.Last1SUIdx = v888; }).Val((Int16)0)));

                if (_.IF(_.EQ(_.NullableNUM(_outer.LangID), (Int16)7)))
                {
                    _outer.Problemtitle = _.CONCAT("=== Bestellbeschreibung ===", " [von Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine);
                    _outer.Diagnosistitle = _.CONCAT("=== Taetigkeitsbeschreibungen ===", VBScriptConstants.vbNewLine);
                    _outer.Solutiontitle = _.CONCAT("=== Loesungsbeschreibung ===", " [von Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine);
                }
                else
                {
                    _outer.Problemtitle = _.CONCAT("=== Orderdescription ===", " [by Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine);
                    _outer.Diagnosistitle = _.CONCAT("=== Diagnosisactivities ===", VBScriptConstants.vbNewLine);
                    _outer.Solutiontitle = _.CONCAT("=== Final solution ===", " [by Agent : ", _outer.Agent1, "]", VBScriptConstants.vbNewLine);
                }
                //Problem-, Diagnose- und Loesungstext auslesen und zusammenfassen
                _outer.Problem = _.VAL(_outer.DescrText);
                _outer.Problem = _.REPLACE(_outer.Problem, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " ");
                if (_.IF(_.NOTEQ(_.NullableSTR(_.LTRIM(_.RTRIM(_outer.Problem))), "")))
                {
                    _outer.ProblemAll = _.CONCAT(_outer.Problemtitle, _outer.Problem, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                }

                //    Dim SUIdx
                var enumerationContent78 = _.ENUMERABLE(_outer.suindices).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent78.MoveNext())
                        break;
                    _outer.SUIdx = enumerationContent78.Current;
                    _outer.SUDiagnosis = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v889 => { _outer.SUIdx = v889; }).Val((Int16)0)));
                    //SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
                    _outer.SURegTime = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v890 => { _outer.SUIdx = v890; }).Val((Int16)0)));
                    _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v891 => { _outer.SUIdx = v891; }).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(_outer.SUDiagnosis), "")))
                    {
                        if (_.IF(_.EQ(_.NullableNUM(_outer.LangID), (Int16)7)))
                        {
                            _outer.SUDiagnosis = "<keine Beschreibung>";
                        }
                        else
                        {
                            _outer.SUDiagnosis = "<no description>";
                        }
                    }
                    _outer.DiagnosisAll = _.CONCAT(_outer.DiagnosisAll, _outer.SUIdx, ". SU (", _outer.Agent, ") -> [", _outer.SURegTime, "]:", VBScriptConstants.vbNewLine, _outer.SUDiagnosis, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                }
                _outer.Solution = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                _outer.Solution = _.REPLACE(_outer.Solution, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " ");
                if (_.IF(_.NOTEQ(_.NullableSTR(_.LTRIM(_.RTRIM(_outer.Solution))), "")))
                {
                    _outer.SolutionAll = _.CONCAT(_outer.SolutionAll, _outer.Solutiontitle, _outer.Solution);
                }
                //Gesammelte Texte in das uebersicht-Textfeld schreiben
                _outer.DescriptionAll = _.CONCAT(_outer.ProblemAll, _outer.Diagnosistitle, _outer.DiagnosisAll, _outer.SolutionAll);
                _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DescriptionAll, v892 => { _outer.DescriptionAll = v892; }));
            }

        }
    }
    public sealed class GlobalReferences : GlobalReferencesBaseT<EnvironmentReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly GlobalReferences _outer;
        private readonly EnvironmentReferences _env;
        public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) : base(compatLayer, env)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = this;
            lcid = null;
            LangID = null;
            hlCase = null;
            Editor = null;
            ActDate = null;
            pCase = null;
            state = null;
            strOrdReqStatus = null;
            hlCaller = null;
            OrderPosIDs = null;
            PosID = null;
            CreateCI = null;
            Counter = null;
            CIisCreated = null;
            CIType = null;
            CIQuantity = null;
            CIQuantityInternal = null;
            ChangedOrderQuantity = null;
            i = null;
            NewCI = null;
            Testname = null;
            OrderNumber = null;
            CompanyCode = null;
            OrderDate = null;
            OrderPosNr = null;
            VendorNumber = null;
            VendorName = null;
            AllocationNumber = null;
            AllocationType = null;
            PlaceOfUnloading = null;
            Incorporation = null;
            PosOrderText = null;
            Reciever = null;
            cn = null;
            rs = null;
            CINumber = null;
            QryString = null;
            Qry = null;
            AssetGroups = null;
            AssetGroup = null;
            AssetGroupID = null;
            rewritten_Group = null;
            ArticleDescription = null;
            CIPrice = null;
            CIPriceUnit = null;
            CIPriceCurrency = null;
            OrderText = null;
            PosOrderInfoText = null;
            CIComment = null;
            DeliveryDate = null;
            objs = null;
            obj = null;
            objtype = null;
            cistatus = null;
            statuscounter = null;
            typecounter = null;
            stornoquantity = null;
            stornocounter = null;
            statusoverview = null;
            CIExistingAtSAPAM = null;
            OrderPosID = null;
            PosType = null;
            suindices = null;
            sumin = null;
            sumindescr = null;
            Agent = null;
            Agent1 = null;
            Last1SUIdx = null;
            LastSU = null;
            DescrText = null;
            DescrTextalt = null;
            DescriptionAll = null;
            ProblemAll = null;
            ProblemAll1 = null;
            DiagnosisAll = null;
            SolutionAll = null;
            Problem = null;
            SUDiagnosis = null;
            SUActivity = null;
            SURegTime = null;
            Solution = null;
            Problemtitle = null;
            Diagnosistitle = null;
            Solutiontitle = null;
            ProblemtitleNew = null;
            SUIdx = null;
        }

        internal object lcid { get; set; }
        internal object LangID { get; set; }
        internal object hlCase { get; set; }
        internal object Editor { get; set; }
        internal object ActDate { get; set; }
        internal object pCase { get; set; }
        internal object state { get; set; }
        internal object strOrdReqStatus { get; set; }
        internal object hlCaller { get; set; }
        internal object OrderPosIDs { get; set; }
        internal object PosID { get; set; }
        internal object CreateCI { get; set; }
        internal object Counter { get; set; }
        internal object CIisCreated { get; set; }
        internal object CIType { get; set; }
        internal object CIQuantity { get; set; }
        internal object CIQuantityInternal { get; set; }
        internal object ChangedOrderQuantity { get; set; }
        internal object i { get; set; }
        internal object NewCI { get; set; }
        internal object Testname { get; set; }
        internal object OrderNumber { get; set; }
        internal object CompanyCode { get; set; }
        internal object OrderDate { get; set; }
        internal object OrderPosNr { get; set; }
        internal object VendorNumber { get; set; }
        internal object VendorName { get; set; }
        internal object AllocationNumber { get; set; }
        internal object AllocationType { get; set; }
        internal object PlaceOfUnloading { get; set; }
        internal object Incorporation { get; set; }
        internal object PosOrderText { get; set; }
        internal object Reciever { get; set; }
        internal object cn { get; set; }
        internal object rs { get; set; }
        internal object CINumber { get; set; }
        internal object QryString { get; set; }
        internal object Qry { get; set; }
        internal object AssetGroups { get; set; }
        internal object AssetGroup { get; set; }
        internal object AssetGroupID { get; set; }
        internal object rewritten_Group { get; set; }
        internal object ArticleDescription { get; set; }
        internal object CIPrice { get; set; }
        internal object CIPriceUnit { get; set; }
        internal object CIPriceCurrency { get; set; }
        internal object OrderText { get; set; }
        internal object PosOrderInfoText { get; set; }
        internal object CIComment { get; set; }
        internal object DeliveryDate { get; set; }
        internal object objs { get; set; }
        internal object obj { get; set; }
        internal object objtype { get; set; }
        internal object cistatus { get; set; }
        internal object statuscounter { get; set; }
        internal object typecounter { get; set; }
        internal object stornoquantity { get; set; }
        internal object stornocounter { get; set; }
        internal object statusoverview { get; set; }
        internal object CIExistingAtSAPAM { get; set; }
        internal object OrderPosID { get; set; }
        internal object PosType { get; set; }
        internal object suindices { get; set; }
        internal object sumin { get; set; }
        internal object sumindescr { get; set; }
        internal object Agent { get; set; }
        internal object Agent1 { get; set; }
        internal object Last1SUIdx { get; set; }
        internal object LastSU { get; set; }
        internal object DescrText { get; set; }
        internal object DescrTextalt { get; set; }
        internal object DescriptionAll { get; set; }
        internal object ProblemAll { get; set; }
        internal object ProblemAll1 { get; set; }
        internal object DiagnosisAll { get; set; }
        internal object SolutionAll { get; set; }
        internal object Problem { get; set; }
        internal object SUDiagnosis { get; set; }
        internal object SUActivity { get; set; }
        internal object SURegTime { get; set; }
        internal object Solution { get; set; }
        internal object Problemtitle { get; set; }
        internal object Diagnosistitle { get; set; }
        internal object Solutiontitle { get; set; }
        internal object ProblemtitleNew { get; set; }
        internal object SUIdx { get; set; }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object ExportObjectIncident { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlITIL2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
