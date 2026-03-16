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
                    _outer.ChangedOrderQuantity = _.VAL(_.CALLm1argp(this, _env.hlITIL2, "CheckIntegerValue", _.ARGS.Val(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v16 => { _outer.PosID = v16; }).Val((Int16)0).Val((Int16)0))).Ref(_env.hlContext, v17 => { _env.hlContext = v17; })));
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
                    _outer.OrderPosNr = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.OrderPosition").Val((Int16)0).Ref(_outer.PosID, v18 => { _outer.PosID = v18; }).Val((Int16)0).Val((Int16)0)));
                    //Abladestelle
                    _outer.PlaceOfUnloading = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PlaceOfUnloading").Val((Int16)0).Ref(_outer.PosID, v19 => { _outer.PosID = v19; }).Val((Int16)0).Val((Int16)0)));
                    //Warenempfaenger
                    _outer.Reciever = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.Reciever").Val((Int16)0).Ref(_outer.PosID, v20 => { _outer.PosID = v20; }).Val((Int16)0).Val((Int16)0)));
                    //Kontierungsnummer
                    _outer.AllocationNumber = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.AllocationNumber").Val((Int16)0).Ref(_outer.PosID, v21 => { _outer.PosID = v21; }).Val((Int16)0).Val((Int16)0)));
                    //LieferDatum
                    _outer.DeliveryDate = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.DeliveryDate").Val((Int16)0).Ref(_outer.PosID, v22 => { _outer.PosID = v22; }).Val((Int16)0).Val((Int16)0)));
                    //Kontierungstyp
                    _outer.AllocationType = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.AllocationType").Val((Int16)0).Ref(_outer.PosID, v23 => { _outer.PosID = v23; }).Val((Int16)0).Val((Int16)0)));
                    //Positionsbestelltext
                    _outer.PosOrderText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PositionOrderText").Val((Int16)0).Ref(_outer.PosID, v24 => { _outer.PosID = v24; }).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(_outer.PosOrderText), "")))
                    {
                        _outer.PosOrderText = " ";
                    }
                    //Positionsinfotext
                    _outer.PosOrderInfoText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PositionInfoNotice").Val((Int16)0).Ref(_outer.PosID, v25 => { _outer.PosID = v25; }).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(_outer.PosOrderInfoText), "")))
                    {
                        _outer.PosOrderInfoText = " ";
                    }
                    //Werk/Standort
                    _outer.Incorporation = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.Incorporation").Val((Int16)0).Ref(_outer.PosID, v26 => { _outer.PosID = v26; }).Val((Int16)0).Val((Int16)0)));
                    //Artikelbeschreibung
                    _outer.ArticleDescription = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ArticleDescription").Val((Int16)0).Ref(_outer.PosID, v27 => { _outer.PosID = v27; }).Val((Int16)0).Val((Int16)0)));
                    //Preis
                    _outer.CIPrice = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Ref(_outer.PosID, v28 => { _outer.PosID = v28; }).Val((Int16)0).Val((Int16)1)));
                    //Preiseinheit
                    _outer.CIPriceCurrency = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Ref(_outer.PosID, v29 => { _outer.PosID = v29; }).Val((Int16)0).Val((Int16)0)));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v30 => { _outer.OrderNumber = v30; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v31 => { _outer.VendorName = v31; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v32 => { _outer.OrderDate = v32; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v33 => { _outer.CompanyCode = v33; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v34 => { _outer.VendorNumber = v34; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v35 => { _outer.OrderPosNr = v35; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v36 => { _outer.AllocationNumber = v36; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v37 => { _outer.AllocationType = v37; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v38 => { _outer.Reciever = v38; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v39 => { _outer.PosID = v39; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v40 => { _outer.PlaceOfUnloading = v40; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v41 => { _outer.CIComment = v41; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v42 => { _outer.ActDate = v42; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v43 => { _outer.DeliveryDate = v43; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v44 => { _outer.ArticleDescription = v44; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v45 => { _outer.CIPrice = v45; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v46 => { _outer.CIPriceCurrency = v46; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v47 => { _outer.AllocationNumber = v47; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v48 => { _outer.AllocationNumber = v48; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v49 => { _outer.CINumber = v49; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v50 => { _outer.NewCI = v50; }));
                                //Neues CI dem Vorgang assoziieren
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v51 => { _outer.hlCase = v51; }).Ref(_outer.NewCI, v52 => { _outer.NewCI = v52; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v53 => { _outer.QryString = v53; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v54 => { _outer.AssetGroup = v54; }).Ref(_outer.NewCI, v55 => { _outer.NewCI = v55; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v56 => { _outer.PosID = v56; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v57 => { _outer.OrderNumber = v57; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v58 => { _outer.VendorName = v58; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v59 => { _outer.OrderDate = v59; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v60 => { _outer.CompanyCode = v60; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v61 => { _outer.VendorNumber = v61; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v62 => { _outer.OrderPosNr = v62; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v63 => { _outer.AllocationNumber = v63; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v64 => { _outer.AllocationType = v64; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v65 => { _outer.Reciever = v65; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v66 => { _outer.PosID = v66; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v67 => { _outer.PlaceOfUnloading = v67; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v68 => { _outer.CIComment = v68; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v69 => { _outer.ActDate = v69; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v70 => { _outer.DeliveryDate = v70; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v71 => { _outer.ArticleDescription = v71; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v72 => { _outer.CIPrice = v72; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v73 => { _outer.CIPriceCurrency = v73; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v74 => { _outer.AllocationNumber = v74; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v75 => { _outer.AllocationNumber = v75; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v76 => { _outer.CINumber = v76; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v77 => { _outer.NewCI = v77; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v78 => { _outer.hlCase = v78; }).Ref(_outer.NewCI, v79 => { _outer.NewCI = v79; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v80 => { _outer.QryString = v80; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v81 => { _outer.AssetGroup = v81; }).Ref(_outer.NewCI, v82 => { _outer.NewCI = v82; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v83 => { _outer.PosID = v83; }).Val((Int16)0).Val("0"));
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
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v84 => { _outer.OrderNumber = v84; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v85 => { _outer.OrderDate = v85; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v86 => { _outer.VendorNumber = v86; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v87 => { _outer.OrderPosNr = v87; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v88 => { _outer.AllocationNumber = v88; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v89 => { _outer.AllocationType = v89; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseStatus.DocumentOrdered").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.SWPlannedAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.SWPlannedDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v90 => { _outer.ActDate = v90; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v91 => { _outer.DeliveryDate = v91; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v92 => { _outer.CIPrice = v92; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v93 => { _outer.CIPriceCurrency = v93; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfSoftwareStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseGeneral.SoftwareLicenseName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v94 => { _outer.ArticleDescription = v94; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIQuantity, v95 => { _outer.CIQuantity = v95; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v96 => { _outer.VendorName = v96; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v97 => { _outer.CompanyCode = v97; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v98 => { _outer.Reciever = v98; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v99 => { _outer.PosID = v99; }));
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v100 => { _outer.PlaceOfUnloading = v100; }));
                        if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                        {
                            _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v101 => { _outer.AllocationNumber = v101; }));
                        }
                        if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                        {
                            _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v102 => { _outer.AllocationNumber = v102; }));
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
                        _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v103 => { _outer.CINumber = v103; }));
                        _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v104 => { _outer.NewCI = v104; }));
                        _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v105 => { _outer.hlCase = v105; }).Ref(_outer.NewCI, v106 => { _outer.NewCI = v106; }).Val(119155));
                        //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                        //Zunaechst ID der Inventargruppe ermitteln
                        _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                        //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                        _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v107 => { _outer.QryString = v107; })));
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
                            _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v108 => { _outer.AssetGroup = v108; }).Ref(_outer.NewCI, v109 => { _outer.NewCI = v109; }).Val(100706));
                        }
                        //Next
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v110 => { _outer.PosID = v110; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v111 => { _outer.OrderNumber = v111; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v112 => { _outer.VendorName = v112; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v113 => { _outer.OrderDate = v113; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v114 => { _outer.CompanyCode = v114; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v115 => { _outer.VendorNumber = v115; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v116 => { _outer.OrderPosNr = v116; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v117 => { _outer.AllocationNumber = v117; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v118 => { _outer.AllocationType = v118; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v119 => { _outer.Reciever = v119; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v120 => { _outer.PosID = v120; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v121 => { _outer.PlaceOfUnloading = v121; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v122 => { _outer.CIComment = v122; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypePrinter"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v123 => { _outer.ActDate = v123; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v124 => { _outer.DeliveryDate = v124; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v125 => { _outer.ArticleDescription = v125; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v126 => { _outer.CIPrice = v126; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v127 => { _outer.CIPriceCurrency = v127; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v128 => { _outer.AllocationNumber = v128; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v129 => { _outer.AllocationNumber = v129; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v130 => { _outer.CINumber = v130; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v131 => { _outer.NewCI = v131; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v132 => { _outer.hlCase = v132; }).Ref(_outer.NewCI, v133 => { _outer.NewCI = v133; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v134 => { _outer.QryString = v134; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v135 => { _outer.AssetGroup = v135; }).Ref(_outer.NewCI, v136 => { _outer.NewCI = v136; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v137 => { _outer.PosID = v137; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v138 => { _outer.OrderNumber = v138; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v139 => { _outer.VendorName = v139; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v140 => { _outer.OrderDate = v140; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v141 => { _outer.CompanyCode = v141; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v142 => { _outer.VendorNumber = v142; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v143 => { _outer.OrderPosNr = v143; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v144 => { _outer.AllocationNumber = v144; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v145 => { _outer.AllocationType = v145; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v146 => { _outer.Reciever = v146; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v147 => { _outer.PosID = v147; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v148 => { _outer.PlaceOfUnloading = v148; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v149 => { _outer.CIComment = v149; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeCopyDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v150 => { _outer.ActDate = v150; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v151 => { _outer.DeliveryDate = v151; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v152 => { _outer.ArticleDescription = v152; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v153 => { _outer.CIPrice = v153; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v154 => { _outer.CIPriceCurrency = v154; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v155 => { _outer.AllocationNumber = v155; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v156 => { _outer.AllocationNumber = v156; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v157 => { _outer.CINumber = v157; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v158 => { _outer.NewCI = v158; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v159 => { _outer.hlCase = v159; }).Ref(_outer.NewCI, v160 => { _outer.NewCI = v160; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v161 => { _outer.QryString = v161; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v162 => { _outer.AssetGroup = v162; }).Ref(_outer.NewCI, v163 => { _outer.NewCI = v163; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v164 => { _outer.PosID = v164; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v165 => { _outer.OrderNumber = v165; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v166 => { _outer.VendorName = v166; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v167 => { _outer.OrderDate = v167; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v168 => { _outer.CompanyCode = v168; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v169 => { _outer.VendorNumber = v169; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v170 => { _outer.OrderPosNr = v170; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v171 => { _outer.AllocationNumber = v171; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v172 => { _outer.AllocationType = v172; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v173 => { _outer.Reciever = v173; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v174 => { _outer.PosID = v174; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v175 => { _outer.PlaceOfUnloading = v175; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v176 => { _outer.CIComment = v176; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeMultiFunctionDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v177 => { _outer.ActDate = v177; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v178 => { _outer.DeliveryDate = v178; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v179 => { _outer.ArticleDescription = v179; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v180 => { _outer.CIPrice = v180; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v181 => { _outer.CIPriceCurrency = v181; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v182 => { _outer.AllocationNumber = v182; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v183 => { _outer.AllocationNumber = v183; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v184 => { _outer.CINumber = v184; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v185 => { _outer.NewCI = v185; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v186 => { _outer.hlCase = v186; }).Ref(_outer.NewCI, v187 => { _outer.NewCI = v187; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v188 => { _outer.QryString = v188; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v189 => { _outer.AssetGroup = v189; }).Ref(_outer.NewCI, v190 => { _outer.NewCI = v190; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v191 => { _outer.PosID = v191; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v192 => { _outer.OrderNumber = v192; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v193 => { _outer.VendorName = v193; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v194 => { _outer.OrderDate = v194; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v195 => { _outer.CompanyCode = v195; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v196 => { _outer.VendorNumber = v196; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v197 => { _outer.OrderPosNr = v197; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v198 => { _outer.AllocationNumber = v198; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v199 => { _outer.AllocationType = v199; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v200 => { _outer.Reciever = v200; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v201 => { _outer.PosID = v201; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v202 => { _outer.PlaceOfUnloading = v202; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v203 => { _outer.CIComment = v203; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("PrintSanDeviceDetail.PrintScanDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("PSDTypeScanner"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v204 => { _outer.ActDate = v204; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v205 => { _outer.DeliveryDate = v205; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v206 => { _outer.ArticleDescription = v206; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v207 => { _outer.CIPrice = v207; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v208 => { _outer.CIPriceCurrency = v208; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v209 => { _outer.AllocationNumber = v209; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v210 => { _outer.AllocationNumber = v210; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v211 => { _outer.CINumber = v211; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v212 => { _outer.NewCI = v212; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v213 => { _outer.hlCase = v213; }).Ref(_outer.NewCI, v214 => { _outer.NewCI = v214; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v215 => { _outer.QryString = v215; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v216 => { _outer.AssetGroup = v216; }).Ref(_outer.NewCI, v217 => { _outer.NewCI = v217; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v218 => { _outer.PosID = v218; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v219 => { _outer.OrderNumber = v219; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v220 => { _outer.VendorName = v220; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v221 => { _outer.OrderDate = v221; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v222 => { _outer.CompanyCode = v222; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v223 => { _outer.VendorNumber = v223; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v224 => { _outer.OrderPosNr = v224; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v225 => { _outer.AllocationNumber = v225; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v226 => { _outer.AllocationType = v226; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v227 => { _outer.Reciever = v227; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v228 => { _outer.PosID = v228; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v229 => { _outer.PlaceOfUnloading = v229; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v230 => { _outer.CIComment = v230; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeMobilePhone"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v231 => { _outer.ActDate = v231; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v232 => { _outer.DeliveryDate = v232; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v233 => { _outer.ArticleDescription = v233; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v234 => { _outer.CIPrice = v234; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v235 => { _outer.CIPriceCurrency = v235; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v236 => { _outer.AllocationNumber = v236; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v237 => { _outer.AllocationNumber = v237; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v238 => { _outer.CINumber = v238; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v239 => { _outer.NewCI = v239; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v240 => { _outer.hlCase = v240; }).Ref(_outer.NewCI, v241 => { _outer.NewCI = v241; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v242 => { _outer.QryString = v242; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v243 => { _outer.AssetGroup = v243; }).Ref(_outer.NewCI, v244 => { _outer.NewCI = v244; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v245 => { _outer.PosID = v245; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v246 => { _outer.OrderNumber = v246; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v247 => { _outer.VendorName = v247; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v248 => { _outer.OrderDate = v248; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v249 => { _outer.CompanyCode = v249; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v250 => { _outer.VendorNumber = v250; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v251 => { _outer.OrderPosNr = v251; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v252 => { _outer.AllocationNumber = v252; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v253 => { _outer.AllocationType = v253; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v254 => { _outer.Reciever = v254; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v255 => { _outer.PosID = v255; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v256 => { _outer.PlaceOfUnloading = v256; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v257 => { _outer.CIComment = v257; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeSIMCard"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v258 => { _outer.ActDate = v258; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v259 => { _outer.DeliveryDate = v259; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v260 => { _outer.ArticleDescription = v260; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v261 => { _outer.CIPrice = v261; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v262 => { _outer.CIPriceCurrency = v262; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v263 => { _outer.AllocationNumber = v263; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v264 => { _outer.AllocationNumber = v264; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v265 => { _outer.CINumber = v265; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v266 => { _outer.NewCI = v266; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v267 => { _outer.hlCase = v267; }).Ref(_outer.NewCI, v268 => { _outer.NewCI = v268; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v269 => { _outer.QryString = v269; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v270 => { _outer.AssetGroup = v270; }).Ref(_outer.NewCI, v271 => { _outer.NewCI = v271; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v272 => { _outer.PosID = v272; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v273 => { _outer.OrderNumber = v273; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v274 => { _outer.VendorName = v274; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v275 => { _outer.OrderDate = v275; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v276 => { _outer.CompanyCode = v276; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v277 => { _outer.VendorNumber = v277; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v278 => { _outer.OrderPosNr = v278; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v279 => { _outer.AllocationNumber = v279; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v280 => { _outer.AllocationType = v280; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v281 => { _outer.Reciever = v281; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v282 => { _outer.PosID = v282; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v283 => { _outer.PlaceOfUnloading = v283; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v284 => { _outer.CIComment = v284; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeUMTSCard"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v285 => { _outer.ActDate = v285; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v286 => { _outer.DeliveryDate = v286; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v287 => { _outer.ArticleDescription = v287; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v288 => { _outer.CIPrice = v288; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v289 => { _outer.CIPriceCurrency = v289; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v290 => { _outer.AllocationNumber = v290; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v291 => { _outer.AllocationNumber = v291; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v292 => { _outer.CINumber = v292; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v293 => { _outer.NewCI = v293; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v294 => { _outer.hlCase = v294; }).Ref(_outer.NewCI, v295 => { _outer.NewCI = v295; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v296 => { _outer.QryString = v296; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v297 => { _outer.AssetGroup = v297; }).Ref(_outer.NewCI, v298 => { _outer.NewCI = v298; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v299 => { _outer.PosID = v299; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v300 => { _outer.OrderNumber = v300; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v301 => { _outer.VendorName = v301; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v302 => { _outer.OrderDate = v302; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v303 => { _outer.CompanyCode = v303; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v304 => { _outer.VendorNumber = v304; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v305 => { _outer.OrderPosNr = v305; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v306 => { _outer.AllocationNumber = v306; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v307 => { _outer.AllocationType = v307; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v308 => { _outer.Reciever = v308; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v309 => { _outer.PosID = v309; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v310 => { _outer.PlaceOfUnloading = v310; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v311 => { _outer.CIComment = v311; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypePDA"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v312 => { _outer.ActDate = v312; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v313 => { _outer.DeliveryDate = v313; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v314 => { _outer.ArticleDescription = v314; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v315 => { _outer.CIPrice = v315; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v316 => { _outer.CIPriceCurrency = v316; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v317 => { _outer.AllocationNumber = v317; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v318 => { _outer.AllocationNumber = v318; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v319 => { _outer.CINumber = v319; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v320 => { _outer.NewCI = v320; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v321 => { _outer.hlCase = v321; }).Ref(_outer.NewCI, v322 => { _outer.NewCI = v322; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v323 => { _outer.QryString = v323; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v324 => { _outer.AssetGroup = v324; }).Ref(_outer.NewCI, v325 => { _outer.NewCI = v325; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v326 => { _outer.PosID = v326; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v327 => { _outer.OrderNumber = v327; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v328 => { _outer.VendorName = v328; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v329 => { _outer.OrderDate = v329; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v330 => { _outer.CompanyCode = v330; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v331 => { _outer.VendorNumber = v331; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v332 => { _outer.OrderPosNr = v332; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v333 => { _outer.AllocationNumber = v333; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v334 => { _outer.AllocationType = v334; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v335 => { _outer.Reciever = v335; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v336 => { _outer.PosID = v336; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v337 => { _outer.PlaceOfUnloading = v337; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v338 => { _outer.CIComment = v338; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MobileDeviceDetail.MobileDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MobileDeviceTypeBlackBerry"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v339 => { _outer.ActDate = v339; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v340 => { _outer.DeliveryDate = v340; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v341 => { _outer.ArticleDescription = v341; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v342 => { _outer.CIPrice = v342; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v343 => { _outer.CIPriceCurrency = v343; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v344 => { _outer.AllocationNumber = v344; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v345 => { _outer.AllocationNumber = v345; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v346 => { _outer.CINumber = v346; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v347 => { _outer.NewCI = v347; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v348 => { _outer.hlCase = v348; }).Ref(_outer.NewCI, v349 => { _outer.NewCI = v349; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v350 => { _outer.QryString = v350; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v351 => { _outer.AssetGroup = v351; }).Ref(_outer.NewCI, v352 => { _outer.NewCI = v352; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v353 => { _outer.PosID = v353; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v354 => { _outer.OrderNumber = v354; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v355 => { _outer.VendorName = v355; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v356 => { _outer.OrderDate = v356; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v357 => { _outer.CompanyCode = v357; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v358 => { _outer.VendorNumber = v358; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v359 => { _outer.OrderPosNr = v359; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v360 => { _outer.AllocationNumber = v360; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v361 => { _outer.AllocationType = v361; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v362 => { _outer.Reciever = v362; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v363 => { _outer.PosID = v363; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v364 => { _outer.PlaceOfUnloading = v364; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v365 => { _outer.CIComment = v365; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v366 => { _outer.ActDate = v366; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v367 => { _outer.DeliveryDate = v367; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v368 => { _outer.ArticleDescription = v368; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v369 => { _outer.CIPrice = v369; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v370 => { _outer.CIPriceCurrency = v370; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v371 => { _outer.AllocationNumber = v371; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v372 => { _outer.AllocationNumber = v372; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v373 => { _outer.CINumber = v373; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v374 => { _outer.NewCI = v374; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v375 => { _outer.hlCase = v375; }).Ref(_outer.NewCI, v376 => { _outer.NewCI = v376; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v377 => { _outer.QryString = v377; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v378 => { _outer.AssetGroup = v378; }).Ref(_outer.NewCI, v379 => { _outer.NewCI = v379; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v380 => { _outer.PosID = v380; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v381 => { _outer.OrderNumber = v381; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v382 => { _outer.VendorName = v382; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v383 => { _outer.OrderDate = v383; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v384 => { _outer.CompanyCode = v384; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v385 => { _outer.VendorNumber = v385; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v386 => { _outer.OrderPosNr = v386; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v387 => { _outer.AllocationNumber = v387; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v388 => { _outer.AllocationType = v388; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v389 => { _outer.Reciever = v389; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v390 => { _outer.PosID = v390; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v391 => { _outer.PlaceOfUnloading = v391; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v392 => { _outer.CIComment = v392; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v393 => { _outer.ActDate = v393; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v394 => { _outer.DeliveryDate = v394; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeBeamer"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v395 => { _outer.ArticleDescription = v395; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v396 => { _outer.CIPrice = v396; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v397 => { _outer.CIPriceCurrency = v397; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v398 => { _outer.AllocationNumber = v398; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v399 => { _outer.AllocationNumber = v399; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v400 => { _outer.CINumber = v400; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v401 => { _outer.NewCI = v401; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v402 => { _outer.hlCase = v402; }).Ref(_outer.NewCI, v403 => { _outer.NewCI = v403; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v404 => { _outer.QryString = v404; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v405 => { _outer.AssetGroup = v405; }).Ref(_outer.NewCI, v406 => { _outer.NewCI = v406; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v407 => { _outer.PosID = v407; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v408 => { _outer.OrderNumber = v408; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v409 => { _outer.VendorName = v409; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v410 => { _outer.OrderDate = v410; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v411 => { _outer.CompanyCode = v411; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v412 => { _outer.VendorNumber = v412; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v413 => { _outer.OrderPosNr = v413; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v414 => { _outer.AllocationNumber = v414; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v415 => { _outer.AllocationType = v415; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v416 => { _outer.Reciever = v416; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v417 => { _outer.PosID = v417; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v418 => { _outer.PlaceOfUnloading = v418; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v419 => { _outer.CIComment = v419; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v420 => { _outer.ActDate = v420; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v421 => { _outer.DeliveryDate = v421; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeVideoConferenceTechnic"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v422 => { _outer.ArticleDescription = v422; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v423 => { _outer.CIPrice = v423; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v424 => { _outer.CIPriceCurrency = v424; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v425 => { _outer.AllocationNumber = v425; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v426 => { _outer.AllocationNumber = v426; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v427 => { _outer.CINumber = v427; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v428 => { _outer.NewCI = v428; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v429 => { _outer.hlCase = v429; }).Ref(_outer.NewCI, v430 => { _outer.NewCI = v430; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v431 => { _outer.QryString = v431; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v432 => { _outer.AssetGroup = v432; }).Ref(_outer.NewCI, v433 => { _outer.NewCI = v433; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v434 => { _outer.PosID = v434; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v435 => { _outer.OrderNumber = v435; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v436 => { _outer.VendorName = v436; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v437 => { _outer.OrderDate = v437; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v438 => { _outer.CompanyCode = v438; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v439 => { _outer.VendorNumber = v439; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v440 => { _outer.OrderPosNr = v440; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v441 => { _outer.AllocationNumber = v441; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v442 => { _outer.AllocationType = v442; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v443 => { _outer.Reciever = v443; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v444 => { _outer.PosID = v444; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v445 => { _outer.PlaceOfUnloading = v445; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v446 => { _outer.CIComment = v446; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v447 => { _outer.ActDate = v447; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v448 => { _outer.DeliveryDate = v448; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("MultiMediaDeviceDetail.MultiMediaDeviceType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("MultiMediaDeviceTypeMediaTechnic"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v449 => { _outer.ArticleDescription = v449; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v450 => { _outer.CIPrice = v450; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v451 => { _outer.CIPriceCurrency = v451; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v452 => { _outer.AllocationNumber = v452; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v453 => { _outer.AllocationNumber = v453; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v454 => { _outer.CINumber = v454; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v455 => { _outer.NewCI = v455; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v456 => { _outer.hlCase = v456; }).Ref(_outer.NewCI, v457 => { _outer.NewCI = v457; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v458 => { _outer.QryString = v458; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v459 => { _outer.AssetGroup = v459; }).Ref(_outer.NewCI, v460 => { _outer.NewCI = v460; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v461 => { _outer.PosID = v461; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v462 => { _outer.OrderNumber = v462; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v463 => { _outer.VendorName = v463; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v464 => { _outer.OrderDate = v464; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v465 => { _outer.CompanyCode = v465; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v466 => { _outer.VendorNumber = v466; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v467 => { _outer.OrderPosNr = v467; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v468 => { _outer.AllocationNumber = v468; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v469 => { _outer.AllocationType = v469; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v470 => { _outer.Reciever = v470; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v471 => { _outer.PosID = v471; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v472 => { _outer.PlaceOfUnloading = v472; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v473 => { _outer.CIComment = v473; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v474 => { _outer.ActDate = v474; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v475 => { _outer.DeliveryDate = v475; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeDictationDevice"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v476 => { _outer.ArticleDescription = v476; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v477 => { _outer.CIPrice = v477; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v478 => { _outer.CIPriceCurrency = v478; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v479 => { _outer.AllocationNumber = v479; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v480 => { _outer.AllocationNumber = v480; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v481 => { _outer.CINumber = v481; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v482 => { _outer.NewCI = v482; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v483 => { _outer.hlCase = v483; }).Ref(_outer.NewCI, v484 => { _outer.NewCI = v484; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v485 => { _outer.QryString = v485; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v486 => { _outer.AssetGroup = v486; }).Ref(_outer.NewCI, v487 => { _outer.NewCI = v487; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v488 => { _outer.PosID = v488; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v489 => { _outer.OrderNumber = v489; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v490 => { _outer.VendorName = v490; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v491 => { _outer.OrderDate = v491; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v492 => { _outer.CompanyCode = v492; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v493 => { _outer.VendorNumber = v493; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v494 => { _outer.OrderPosNr = v494; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v495 => { _outer.AllocationNumber = v495; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v496 => { _outer.AllocationType = v496; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v497 => { _outer.Reciever = v497; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v498 => { _outer.PosID = v498; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v499 => { _outer.PlaceOfUnloading = v499; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v500 => { _outer.CIComment = v500; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v501 => { _outer.ActDate = v501; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v502 => { _outer.DeliveryDate = v502; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeUSV"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v503 => { _outer.ArticleDescription = v503; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v504 => { _outer.CIPrice = v504; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v505 => { _outer.CIPriceCurrency = v505; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v506 => { _outer.AllocationNumber = v506; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v507 => { _outer.AllocationNumber = v507; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v508 => { _outer.CINumber = v508; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v509 => { _outer.NewCI = v509; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v510 => { _outer.hlCase = v510; }).Ref(_outer.NewCI, v511 => { _outer.NewCI = v511; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v512 => { _outer.QryString = v512; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v513 => { _outer.AssetGroup = v513; }).Ref(_outer.NewCI, v514 => { _outer.NewCI = v514; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v515 => { _outer.PosID = v515; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v516 => { _outer.OrderNumber = v516; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v517 => { _outer.VendorName = v517; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v518 => { _outer.OrderDate = v518; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v519 => { _outer.CompanyCode = v519; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v520 => { _outer.VendorNumber = v520; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v521 => { _outer.OrderPosNr = v521; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v522 => { _outer.AllocationNumber = v522; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v523 => { _outer.AllocationType = v523; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v524 => { _outer.Reciever = v524; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v525 => { _outer.PosID = v525; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v526 => { _outer.PlaceOfUnloading = v526; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v527 => { _outer.CIComment = v527; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v528 => { _outer.ActDate = v528; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v529 => { _outer.DeliveryDate = v529; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeControlCam"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v530 => { _outer.ArticleDescription = v530; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v531 => { _outer.CIPrice = v531; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v532 => { _outer.CIPriceCurrency = v532; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v533 => { _outer.AllocationNumber = v533; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v534 => { _outer.AllocationNumber = v534; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v535 => { _outer.CINumber = v535; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v536 => { _outer.NewCI = v536; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v537 => { _outer.hlCase = v537; }).Ref(_outer.NewCI, v538 => { _outer.NewCI = v538; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v539 => { _outer.QryString = v539; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v540 => { _outer.AssetGroup = v540; }).Ref(_outer.NewCI, v541 => { _outer.NewCI = v541; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v542 => { _outer.PosID = v542; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v543 => { _outer.OrderNumber = v543; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v544 => { _outer.VendorName = v544; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v545 => { _outer.OrderDate = v545; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v546 => { _outer.CompanyCode = v546; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v547 => { _outer.VendorNumber = v547; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v548 => { _outer.OrderPosNr = v548; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v549 => { _outer.AllocationNumber = v549; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v550 => { _outer.AllocationType = v550; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v551 => { _outer.Reciever = v551; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v552 => { _outer.PosID = v552; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v553 => { _outer.PlaceOfUnloading = v553; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v554 => { _outer.CIComment = v554; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v555 => { _outer.ActDate = v555; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v556 => { _outer.DeliveryDate = v556; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeBDE"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v557 => { _outer.ArticleDescription = v557; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v558 => { _outer.CIPrice = v558; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v559 => { _outer.CIPriceCurrency = v559; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v560 => { _outer.AllocationNumber = v560; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v561 => { _outer.AllocationNumber = v561; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v562 => { _outer.CINumber = v562; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v563 => { _outer.NewCI = v563; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v564 => { _outer.hlCase = v564; }).Ref(_outer.NewCI, v565 => { _outer.NewCI = v565; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v566 => { _outer.QryString = v566; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v567 => { _outer.AssetGroup = v567; }).Ref(_outer.NewCI, v568 => { _outer.NewCI = v568; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v569 => { _outer.PosID = v569; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v570 => { _outer.OrderNumber = v570; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v571 => { _outer.VendorName = v571; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v572 => { _outer.OrderDate = v572; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v573 => { _outer.CompanyCode = v573; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v574 => { _outer.VendorNumber = v574; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v575 => { _outer.OrderPosNr = v575; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v576 => { _outer.AllocationNumber = v576; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v577 => { _outer.AllocationType = v577; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v578 => { _outer.Reciever = v578; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v579 => { _outer.PosID = v579; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v580 => { _outer.PlaceOfUnloading = v580; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v581 => { _outer.CIComment = v581; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v582 => { _outer.ActDate = v582; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v583 => { _outer.DeliveryDate = v583; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeSpaceMouse"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v584 => { _outer.ArticleDescription = v584; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v585 => { _outer.CIPrice = v585; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v586 => { _outer.CIPriceCurrency = v586; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v587 => { _outer.AllocationNumber = v587; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v588 => { _outer.AllocationNumber = v588; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v589 => { _outer.CINumber = v589; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v590 => { _outer.NewCI = v590; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v591 => { _outer.hlCase = v591; }).Ref(_outer.NewCI, v592 => { _outer.NewCI = v592; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v593 => { _outer.QryString = v593; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v594 => { _outer.AssetGroup = v594; }).Ref(_outer.NewCI, v595 => { _outer.NewCI = v595; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v596 => { _outer.PosID = v596; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v597 => { _outer.OrderNumber = v597; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v598 => { _outer.VendorName = v598; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v599 => { _outer.OrderDate = v599; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v600 => { _outer.CompanyCode = v600; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v601 => { _outer.VendorNumber = v601; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v602 => { _outer.OrderPosNr = v602; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v603 => { _outer.AllocationNumber = v603; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v604 => { _outer.AllocationType = v604; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v605 => { _outer.Reciever = v605; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v606 => { _outer.PosID = v606; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v607 => { _outer.PlaceOfUnloading = v607; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v608 => { _outer.CIComment = v608; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v609 => { _outer.ActDate = v609; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v610 => { _outer.DeliveryDate = v610; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("NetworkComponentDetail.NetworkComponentType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("TypeActiveNetworkComponet"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v611 => { _outer.ArticleDescription = v611; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v612 => { _outer.CIPrice = v612; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v613 => { _outer.CIPriceCurrency = v613; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v614 => { _outer.AllocationNumber = v614; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v615 => { _outer.AllocationNumber = v615; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v616 => { _outer.CINumber = v616; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v617 => { _outer.NewCI = v617; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v618 => { _outer.hlCase = v618; }).Ref(_outer.NewCI, v619 => { _outer.NewCI = v619; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v620 => { _outer.QryString = v620; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v621 => { _outer.AssetGroup = v621; }).Ref(_outer.NewCI, v622 => { _outer.NewCI = v622; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v623 => { _outer.PosID = v623; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v624 => { _outer.OrderNumber = v624; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v625 => { _outer.VendorName = v625; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v626 => { _outer.OrderDate = v626; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v627 => { _outer.CompanyCode = v627; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v628 => { _outer.VendorNumber = v628; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v629 => { _outer.OrderPosNr = v629; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v630 => { _outer.AllocationNumber = v630; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v631 => { _outer.AllocationType = v631; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v632 => { _outer.Reciever = v632; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v633 => { _outer.PosID = v633; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v634 => { _outer.PlaceOfUnloading = v634; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v635 => { _outer.CIComment = v635; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v636 => { _outer.ActDate = v636; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v637 => { _outer.DeliveryDate = v637; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("NetworkComponentDetail.NetworkComponentType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("TypeHomeOfficeRouter"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v638 => { _outer.ArticleDescription = v638; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v639 => { _outer.CIPrice = v639; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v640 => { _outer.CIPriceCurrency = v640; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v641 => { _outer.AllocationNumber = v641; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v642 => { _outer.AllocationNumber = v642; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v643 => { _outer.CINumber = v643; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v644 => { _outer.NewCI = v644; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v645 => { _outer.hlCase = v645; }).Ref(_outer.NewCI, v646 => { _outer.NewCI = v646; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v647 => { _outer.QryString = v647; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v648 => { _outer.AssetGroup = v648; }).Ref(_outer.NewCI, v649 => { _outer.NewCI = v649; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v650 => { _outer.PosID = v650; }).Val((Int16)0).Val("0"));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v651 => { _outer.OrderNumber = v651; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v652 => { _outer.VendorName = v652; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v653 => { _outer.OrderDate = v653; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v654 => { _outer.CompanyCode = v654; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v655 => { _outer.VendorNumber = v655; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v656 => { _outer.OrderPosNr = v656; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v657 => { _outer.AllocationNumber = v657; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v658 => { _outer.AllocationType = v658; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v659 => { _outer.Reciever = v659; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v660 => { _outer.PosID = v660; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v661 => { _outer.PlaceOfUnloading = v661; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v662 => { _outer.CIComment = v662; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v663 => { _outer.ActDate = v663; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v664 => { _outer.DeliveryDate = v664; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeHeadset"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v665 => { _outer.ArticleDescription = v665; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v666 => { _outer.CIPrice = v666; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v667 => { _outer.CIPriceCurrency = v667; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v668 => { _outer.AllocationNumber = v668; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v669 => { _outer.AllocationNumber = v669; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v670 => { _outer.CINumber = v670; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v671 => { _outer.NewCI = v671; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v672 => { _outer.hlCase = v672; }).Ref(_outer.NewCI, v673 => { _outer.NewCI = v673; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v674 => { _outer.QryString = v674; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v675 => { _outer.AssetGroup = v675; }).Ref(_outer.NewCI, v676 => { _outer.NewCI = v676; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v677 => { _outer.PosID = v677; }).Val((Int16)0).Val("0"));

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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v678 => { _outer.OrderNumber = v678; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v679 => { _outer.VendorName = v679; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v680 => { _outer.OrderDate = v680; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v681 => { _outer.CompanyCode = v681; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v682 => { _outer.VendorNumber = v682; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v683 => { _outer.OrderPosNr = v683; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v684 => { _outer.AllocationNumber = v684; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v685 => { _outer.AllocationType = v685; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v686 => { _outer.Reciever = v686; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v687 => { _outer.PosID = v687; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v688 => { _outer.PlaceOfUnloading = v688; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v689 => { _outer.CIComment = v689; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v690 => { _outer.ActDate = v690; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v691 => { _outer.DeliveryDate = v691; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("GenericAssetDetail.GenericAssetType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("GenericAssetTypeConferencePhone"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v692 => { _outer.ArticleDescription = v692; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v693 => { _outer.CIPrice = v693; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v694 => { _outer.CIPriceCurrency = v694; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v695 => { _outer.AllocationNumber = v695; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v696 => { _outer.AllocationNumber = v696; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v697 => { _outer.CINumber = v697; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v698 => { _outer.NewCI = v698; }));
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v699 => { _outer.hlCase = v699; }).Ref(_outer.NewCI, v700 => { _outer.NewCI = v700; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                //hlContext.Trace 1, "Suche Inv-Gruppe: " & QryString
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v701 => { _outer.QryString = v701; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v702 => { _outer.AssetGroup = v702; }).Ref(_outer.NewCI, v703 => { _outer.NewCI = v703; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v704 => { _outer.PosID = v704; }).Val((Int16)0).Val("0"));

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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderNumber, v705 => { _outer.OrderNumber = v705; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.VendorName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorName, v706 => { _outer.VendorName = v706; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderDate, v707 => { _outer.OrderDate = v707; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CompanyCode, v708 => { _outer.CompanyCode = v708; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.VendorNumber, v709 => { _outer.VendorNumber = v709; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.OrderPosNr, v710 => { _outer.OrderPosNr = v710; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v711 => { _outer.AllocationNumber = v711; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationType, v712 => { _outer.AllocationType = v712; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.GoodsRecipient").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.Reciever, v713 => { _outer.Reciever = v713; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.OrderPosID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PosID, v714 => { _outer.PosID = v714; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.PlaceOfUnloading").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.PlaceOfUnloading, v715 => { _outer.PlaceOfUnloading = v715; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.LongComment").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIComment, v716 => { _outer.CIComment = v716; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrder").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderAgent").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("System"));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIOrderDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v717 => { _outer.ActDate = v717; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("ProcurementDetail.DeliveryDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DeliveryDate, v718 => { _outer.DeliveryDate = v718; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ArticleDescription, v719 => { _outer.ArticleDescription = v719; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_VALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPrice, v720 => { _outer.CIPrice = v720; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.PurchasePrice.CURRENCY_SYMBOL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CIPriceCurrency, v721 => { _outer.CIPriceCurrency = v721; }));
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CONCAT(_.ADD(_.ADD("Beschaffung/Order am/at: ", _outer.ActDate), " von/by: System"), VBScriptConstants.vbNewLine)));
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "K")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v722 => { _outer.AllocationNumber = v722; }));
                                }
                                if (_.IF(_.EQ(_.NullableSTR(_outer.AllocationType), "A")))
                                {
                                    _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.AllocationNumber, v723 => { _outer.AllocationNumber = v723; }));
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
                                _.CALLm1argp(this, _outer.NewCI, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.CINumber, v724 => { _outer.CINumber = v724; }));
                                _.CALLm1argp(this, _env.hlContext, "saveobject", _.ARGS.Ref(_outer.NewCI, v725 => { _outer.NewCI = v725; }));
                                //Neues CI dem Vorgang assoziieren
                                _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.hlCase, v726 => { _outer.hlCase = v726; }).Ref(_outer.NewCI, v727 => { _outer.NewCI = v727; }).Val(119155));
                                //Neues CI der Abladestelle (namensgleiche Inventargruppe) assoziieren
                                //Zunaechst ID der Inventargruppe ermitteln
                                _outer.QryString = _.CONCAT("SEARCH AssetGroup WHERE AssetGroupGeneral.AssetGroupName = ", _outer.Incorporation);
                                _.CALLm1v2(this, _env.hlContext, "Trace", (Int16)1, _.CONCAT("Suche Inv-Gruppe: ", _outer.QryString));
                                _outer.Qry = _.OBJ(_.CALLm1argp(this, _env.hlContext, "OpenSearch", _.ARGS.Ref(_outer.QryString, v728 => { _outer.QryString = v728; })));
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
                                    _.CALLm1argp(this, _env.hlContext, "CreateAssociation", _.ARGS.Ref(_outer.AssetGroup, v729 => { _outer.AssetGroup = v729; }).Ref(_outer.NewCI, v730 => { _outer.NewCI = v730; }).Val(100706));
                                }
                            }
                        }
                        //Geaenderte Bestellmenge der alten Bestellmenge dazu addieren und anschliessend die geaenderte Bestellmenge auf 0 setzen
                        //CIQuantityInternal = CIQuantityInternal + ChangedOrderQuantity
                        //hlContext.Trace 1, "Interne Bestellmenge: " & CIQuantityInternal
                        //hlcase.SetValue "OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity",0,PosID,0,CIQuantityInternal
                        //Geaenderte Bestellmenge auf 0 setzen
                        _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.ChangedOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v731 => { _outer.PosID = v731; }).Val((Int16)0).Val("0"));
                    }
                    //Kennzeichnen, dass CI erzeugt wurde
                    _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIisCreated").Val((Int16)0).Ref(_outer.PosID, v732 => { _outer.PosID = v732; }).Val((Int16)0).Val("1"));
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
                _outer.CreateCI = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CreateCI").Val((Int16)0).Ref(_outer.PosID, v733 => { _outer.PosID = v733; }).Val((Int16)0).Val((Int16)0)));
                //Pruefen ob CI bereits erzeugt wurde
                _outer.CIisCreated = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIisCreated").Val((Int16)0).Ref(_outer.PosID, v734 => { _outer.PosID = v734; }).Val((Int16)0).Val((Int16)0)));
                //Geraetetyp validieren
                _outer.PosType = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.CIType").Val((Int16)0).Ref(_outer.PosID, v735 => { _outer.PosID = v735; }).Val((Int16)0).Val((Int16)0)));
                //Bestellmenge abfragen
                _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v736 => { _outer.PosID = v736; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v737 => { _outer.ActDate = v737; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v738 => { _outer.statusoverview = v738; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v739 => { _env.hlContext = v739; }).Ref(_outer.obj, v740 => { _outer.obj = v740; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v741 => { _outer.obj = v741; }));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v742 => { _outer.ActDate = v742; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v743 => { _outer.statusoverview = v743; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v744 => { _env.hlContext = v744; }).Ref(_outer.obj, v745 => { _outer.obj = v745; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v746 => { _outer.obj = v746; }));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v747 => { _outer.ActDate = v747; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v748 => { _outer.statusoverview = v748; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v749 => { _env.hlContext = v749; }).Ref(_outer.obj, v750 => { _outer.obj = v750; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v751 => { _outer.obj = v751; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v752 => { _outer.PosID = v752; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v753 => { _outer.ActDate = v753; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v754 => { _outer.statusoverview = v754; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v755 => { _env.hlContext = v755; }).Ref(_outer.obj, v756 => { _outer.obj = v756; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v757 => { _outer.obj = v757; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v758 => { _outer.PosID = v758; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v759 => { _outer.ActDate = v759; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v760 => { _outer.statusoverview = v760; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v761 => { _env.hlContext = v761; }).Ref(_outer.obj, v762 => { _outer.obj = v762; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v763 => { _outer.obj = v763; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v764 => { _outer.PosID = v764; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v765 => { _outer.ActDate = v765; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v766 => { _outer.statusoverview = v766; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v767 => { _env.hlContext = v767; }).Ref(_outer.obj, v768 => { _outer.obj = v768; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v769 => { _outer.obj = v769; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v770 => { _outer.PosID = v770; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v771 => { _outer.ActDate = v771; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v772 => { _outer.statusoverview = v772; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v773 => { _env.hlContext = v773; }).Ref(_outer.obj, v774 => { _outer.obj = v774; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v775 => { _outer.obj = v775; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v776 => { _outer.PosID = v776; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v777 => { _outer.ActDate = v777; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v778 => { _outer.statusoverview = v778; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v779 => { _env.hlContext = v779; }).Ref(_outer.obj, v780 => { _outer.obj = v780; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v781 => { _outer.obj = v781; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v782 => { _outer.PosID = v782; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v783 => { _outer.ActDate = v783; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v784 => { _outer.statusoverview = v784; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v785 => { _env.hlContext = v785; }).Ref(_outer.obj, v786 => { _outer.obj = v786; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v787 => { _outer.obj = v787; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v788 => { _outer.PosID = v788; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v789 => { _outer.ActDate = v789; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v790 => { _outer.statusoverview = v790; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v791 => { _env.hlContext = v791; }).Ref(_outer.obj, v792 => { _outer.obj = v792; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v793 => { _outer.obj = v793; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v794 => { _outer.PosID = v794; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v795 => { _outer.ActDate = v795; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v796 => { _outer.statusoverview = v796; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v797 => { _env.hlContext = v797; }).Ref(_outer.obj, v798 => { _outer.obj = v798; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v799 => { _outer.obj = v799; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v800 => { _outer.PosID = v800; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v801 => { _outer.ActDate = v801; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v802 => { _outer.statusoverview = v802; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v803 => { _env.hlContext = v803; }).Ref(_outer.obj, v804 => { _outer.obj = v804; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v805 => { _outer.obj = v805; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v806 => { _outer.PosID = v806; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v807 => { _outer.ActDate = v807; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v808 => { _outer.statusoverview = v808; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v809 => { _env.hlContext = v809; }).Ref(_outer.obj, v810 => { _outer.obj = v810; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v811 => { _outer.obj = v811; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v812 => { _outer.PosID = v812; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v813 => { _outer.ActDate = v813; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v814 => { _outer.statusoverview = v814; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v815 => { _env.hlContext = v815; }).Ref(_outer.obj, v816 => { _outer.obj = v816; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v817 => { _outer.obj = v817; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v818 => { _outer.PosID = v818; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v819 => { _outer.ActDate = v819; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v820 => { _outer.statusoverview = v820; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v821 => { _env.hlContext = v821; }).Ref(_outer.obj, v822 => { _outer.obj = v822; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v823 => { _outer.obj = v823; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v824 => { _outer.PosID = v824; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v825 => { _outer.ActDate = v825; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v826 => { _outer.statusoverview = v826; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v827 => { _env.hlContext = v827; }).Ref(_outer.obj, v828 => { _outer.obj = v828; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v829 => { _outer.obj = v829; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v830 => { _outer.PosID = v830; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v831 => { _outer.ActDate = v831; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v832 => { _outer.statusoverview = v832; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v833 => { _env.hlContext = v833; }).Ref(_outer.obj, v834 => { _outer.obj = v834; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v835 => { _outer.obj = v835; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v836 => { _outer.PosID = v836; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v837 => { _outer.ActDate = v837; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v838 => { _outer.statusoverview = v838; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v839 => { _env.hlContext = v839; }).Ref(_outer.obj, v840 => { _outer.obj = v840; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v841 => { _outer.obj = v841; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v842 => { _outer.PosID = v842; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v843 => { _outer.ActDate = v843; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v844 => { _outer.statusoverview = v844; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v845 => { _env.hlContext = v845; }).Ref(_outer.obj, v846 => { _outer.obj = v846; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v847 => { _outer.obj = v847; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v848 => { _outer.PosID = v848; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v849 => { _outer.ActDate = v849; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v850 => { _outer.statusoverview = v850; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v851 => { _env.hlContext = v851; }).Ref(_outer.obj, v852 => { _outer.obj = v852; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v853 => { _outer.obj = v853; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v854 => { _outer.PosID = v854; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v855 => { _outer.ActDate = v855; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v856 => { _outer.statusoverview = v856; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v857 => { _env.hlContext = v857; }).Ref(_outer.obj, v858 => { _outer.obj = v858; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v859 => { _outer.obj = v859; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v860 => { _outer.PosID = v860; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v861 => { _outer.ActDate = v861; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v862 => { _outer.statusoverview = v862; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v863 => { _env.hlContext = v863; }).Ref(_outer.obj, v864 => { _outer.obj = v864; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v865 => { _outer.obj = v865; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v866 => { _outer.PosID = v866; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v867 => { _outer.ActDate = v867; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v868 => { _outer.statusoverview = v868; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v869 => { _env.hlContext = v869; }).Ref(_outer.obj, v870 => { _outer.obj = v870; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v871 => { _outer.obj = v871; }));
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
                    _outer.CIQuantity = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestAttribute.OrderedCIs_CA.PurchaseOrderQuantity").Val((Int16)0).Ref(_outer.PosID, v872 => { _outer.PosID = v872; }).Val((Int16)0).Val((Int16)0)));
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
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIEliminationDate").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.ActDate, v873 => { _outer.ActDate = v873; }));
                                            _outer.statusoverview = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            _outer.statusoverview = _.CONCAT(_outer.statusoverview, _.ADD(_.ADD(_.ADD(VBScriptConstants.vbNewLine, "Eliminierung/Elimination am/at: "), _outer.ActDate), " durch/by: System"), VBScriptConstants.vbNewLine);
                                            _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.CIStatusOverview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.statusoverview, v874 => { _outer.statusoverview = v874; }));
                                            //Incident erzeugen, wenn CI bereits in SAP AM angelegt wurde
                                            _outer.CIExistingAtSAPAM = _.VAL(_.CALLm1argp(this, _outer.obj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CIExistingAtSAPAM").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                                            if (_.IF(_.EQ(_.NullableSTR(_outer.CIExistingAtSAPAM), "1")))
                                            {
                                                _.CALLm0argp(this, _env.ExportObjectIncident, _.ARGS.Ref(_env.hlContext, v875 => { _env.hlContext = v875; }).Ref(_outer.obj, v876 => { _outer.obj = v876; }));
                                                _.CALLm1argp(this, _outer.obj, "SetValue", _.ARGS.Val("TrumpfAssetStatus.IncidentBecauseOfCIElimination").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                                            }
                                            _.CALLm1argp(this, _env.hlContext, "SaveObject", _.ARGS.Ref(_outer.obj, v877 => { _outer.obj = v877; }));
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
            _outer.Last1SUIdx = _.VAL(_.CALLm1argp(this, _env.hlITIL2, "GetLastSUIdx", _.ARGS.Ref(_outer.hlCase, v878 => { _outer.hlCase = v878; }).Ref(_env.hlContext, v879 => { _env.hlContext = v879; })));
            //Index vorletzte SU
            _outer.LastSU = _.SUBT(_outer.Last1SUIdx, (Int16)1);
            _outer.DescrText = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            //Urspruenlichen Beschreibungstext ermitteln
            _outer.sumindescr = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestSUAttribute.CaseDescriptionSU").Val((Int16)0).Val((Int16)0).Ref(_outer.sumin, v880 => { _outer.sumin = v880; }).Val((Int16)0)));
            _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v881 => { _outer.Last1SUIdx = v881; }).Val((Int16)0)));
            _outer.Agent1 = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.sumin, v882 => { _outer.sumin = v882; }).Val((Int16)0)));

            //----------------------------------------------------------------------------------------------------------
            //Kumuliert die Texte der Bearbeitungsschritte und schreibt sie in das
            //Overview-Textfeld. Die Texte werden durch Trennzeichen voneinander abgegrenzt.
            //Pruefen ob mehr als 1 SU
            _outer.DescrTextalt = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("OrderRequestSUAttribute.CaseDescriptionSU").Val((Int16)0).Val((Int16)0).Ref(_outer.LastSU, v883 => { _outer.LastSU = v883; }).Val((Int16)0)));
            if (_.IF(_.GT(_.NullableNUM(_outer.LastSU), (Int16)0)))
            {
                //Pruefen, ob Beschreibungstext sich geaendert hat
                if (_.IF(_.NOTEQ(_outer.DescrText, _outer.sumindescr)))
                {
                    _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v884 => { _outer.Last1SUIdx = v884; }).Val((Int16)0)));
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
                        _outer.SUDiagnosis = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v885 => { _outer.SUIdx = v885; }).Val((Int16)0)));
                        //SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
                        _outer.SURegTime = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v886 => { _outer.SUIdx = v886; }).Val((Int16)0)));
                        _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v887 => { _outer.SUIdx = v887; }).Val((Int16)0)));
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
                    _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DescriptionAll, v888 => { _outer.DescriptionAll = v888; }));
                }
            }
            if (_.IF(_.EQ(_outer.sumindescr, _outer.DescrText)))
            {
                _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.Last1SUIdx, v889 => { _outer.Last1SUIdx = v889; }).Val((Int16)0)));

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
                    _outer.SUDiagnosis = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v890 => { _outer.SUIdx = v890; }).Val((Int16)0)));
                    //SUActivity = hlcase.GetValue("IncidentSUAttribute.IncidentOperation", LangID, 0, SUIdx, 0)
                    _outer.SURegTime = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v891 => { _outer.SUIdx = v891; }).Val((Int16)0)));
                    _outer.Agent = _.VAL(_.CALLm1argp(this, _outer.hlCase, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(_outer.SUIdx, v892 => { _outer.SUIdx = v892; }).Val((Int16)0)));
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
                _.CALLm1argp(this, _outer.hlCase, "SetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(_outer.DescriptionAll, v893 => { _outer.DescriptionAll = v893; }));
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
