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
        }

        public void IncReqOnLoad()
        {
            object rewritten_ReadOnly = null;
            object NoPerson = null;
            object NoAsset = null;
            object Anfrageart = null;
            object Valid = null; /* Undeclared in source */
            object VIP = null; /* Undeclared in source */
            object hlProduct = null; /* Undeclared in source */
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            object varString = null; /* Undeclared in source */
            object varAType = null; /* Undeclared in source */
            rewritten_ReadOnly = true;
            NoPerson = true;
            NoAsset = true;

            //Zunächst überprüfen ob der Vorgang schreibgeschützt ist
            //First of all check whether the Case is write protected
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.hlobj, "IsReadOnly", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0))), (Int16)0)))
            {
                rewritten_ReadOnly = false;
            }

            //Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
            //Check wether the Caller object exist
            if (_.IF(_.AND(_.EQ(_.ISOBJECT(_env.hlcaller), true), _.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditSurname, "Text")), ""))))
            {
                NoPerson = false;
            }

            //VIP-Status des Anfragers abfragen und im Vorgang setzen
            Valid = _.VAL(_.CALL(this, _env.hlcaller, "HasContent", _.ARGS.Val("PersonGeneral.VIPLevel").Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableNUM(Valid), (Int16)1)))
            {
                VIP = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.VIPLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                //If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
                if (_.IF(_.EQ(VIP, "VIPLevelVIP")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)1));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)142, (Int16)139, (Int16)254)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelITAdminDitzingen")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)2));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelITAdminTG")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)3));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelSAPKeyUserTUS")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)4));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelNon")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)0));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET("", this, _env.Person, "BackColor");
                }
            }

            //Prüft ob ein Produkt Objekt vorhanden ist und ob dieses auch angezeigt wird
            //Check wether the Product object exist
            if (_.IF(_.AND(_.EQ(_.ISOBJECT(hlProduct), true), _.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditAssetModel, "Text")), ""))))
            {
                NoAsset = false;
            }

            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v => { lcid = v; })));

            //Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check requester search status to set the caption of the button
            if (_.IF(_.EQ(NoPerson, false)))
            {
                if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchCaller, "GetSearchState")), (Int16)3)))
                {
                    _.SET("Reset", this, _env.SearchCaller, "Caption");
                }
                else
                {
                    _.SET("Betroffener", this, _env.SearchCaller, "Caption");
                }
            }

            //Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check Asset search status to set the caption of the button
            if (_.IF(_.EQ(NoAsset, false)))
            {
                if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchAsset, "GetSearchState")), (Int16)3)))
                {
                    _.SET("Reset", this, _env.SearchAsset, "Caption");
                }
                else
                {
                    _.SET("Inventar", this, _env.SearchAsset, "Caption");
                }
            }

            if (_.IF(_.EQ(NoAsset, false)))
            {
                //Setzen des Inventars
                //Setting the asset
                varString = "";
                varAType = _.VAL(_.CALL(this, hlProduct, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.OR(_.OR(_.OR(_.EQ(_.NullableSTR(varAType), "DesktopComputer"), _.EQ(_.NullableSTR(varAType), "ServerComputer")), _.EQ(_.NullableSTR(varAType), "NotebookComputer")), _.EQ(_.NullableSTR(varAType), "Printer"))))
                {
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditHostname, "Text")), "")))
                    {
                        varString = _.VAL(_.CALL(this, _env.EditHostname, "Text"));
                    }
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditAssetModel, "Text")), "")))
                    {
                        varString = _.CONCAT(varString, " ", _.CALL(this, _env.EditAssetModel, "Text"));
                    }
                }
                else
                {
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditAssetModel, "Text")), "")))
                    {
                        varString = _.VAL(_.CALL(this, _env.EditAssetModel, "Text"));
                    }
                    else
                    {
                        _.SET(" ", this, _env.EditAssetModel, "Text");
                    }
                }
                _.SET(_.VAL(varString), this, _env.EditAssetModel, "Text");
            }

            //Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
            Anfrageart = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeIncident")))
            {
                _.SET(true, this, _env.ComboImpact, "Disabled");
                _.SET(true, this, _env.ComboFunctionalRange, "Disabled");
            }
            else
            {
                _.SET(false, this, _env.ComboImpact, "Disabled");
                _.SET(false, this, _env.ComboFunctionalRange, "Disabled");
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeContact")))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");

            }
            else
            {
                _.SET(true, this, _env.CaseProblem, "Disabled");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(true, this, _env.CaseDiagnosis, "Disabled");
                _.SET(true, this, _env.KeywordTree, "Disabled");
                _.SET(true, this, _env.Attachment, "Disabled");
                _.SET(true, this, _env.ComboProductionalRelevanz, "Disabled");
                _.SET(true, this, _env.ComboIncidentStatus, "Disabled");
            }

            //Zugriff auf Übersichts-Buttons regeln
            if (_.IF(_.EQ(rewritten_ReadOnly, false)))
            {
                _.SET(false, this, _env.ButtonShowOverView, "Disabled");
                _.SET(false, this, _env.ButtonEmailPreview, "Disabled");
                _.SET(false, this, _env.EditSubjectCase, "Disabled");
            }
            else
            {
                _.SET(true, this, _env.ButtonShowOverView, "Disabled");
                _.SET(true, this, _env.ButtonEmailPreview, "Disabled");
                _.SET(true, this, _env.EditSubjectCase, "Disabled");
            }

            //Einfärben der GrupBox CaseAttributes je nach Priorität
            object target = _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseClassificationAttribute.Priority").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0));
            if (_.IF(_.EQ(target, "Priority1")))
            {
                _.SET(_.VAL(_.RGB((Int16)107, (Int16)105, (Int16)248)), this, _env.CaseAttributes, "BackColor");
            }
            else if (_.IF(_.EQ(target, "Priority2")))
            {
                _.SET(_.VAL(_.RGB((Int16)119, (Int16)170, (Int16)251)), this, _env.CaseAttributes, "BackColor");
            }
            else if (_.IF(_.EQ(target, "Priority3")))
            {
                _.SET(_.VAL(_.RGB((Int16)132, (Int16)235, (Int16)255)), this, _env.CaseAttributes, "BackColor");
            }
            else if (_.IF(_.EQ(target, "Priority4")))
            {
                _.SET(_.VAL(_.RGB((Int16)128, (Int16)213, (Int16)177)), this, _env.CaseAttributes, "BackColor");
            }
            else if (_.IF(_.EQ(target, "Priority5")))
            {
                _.SET(_.VAL(_.RGB((Int16)123, (Int16)190, (Int16)99)), this, _env.CaseAttributes, "BackColor");
            }
            else
            {
                _.SET(_.VAL(_.RGB((Int16)248, (Int16)245, (Int16)240)), this, _env.CaseAttributes, "BackColor");
            }

            //Bei Status ToProof wird die Email-Tab angewählt
            if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))), "IncidentStatusToProof")))
            {
                _.SET(true, this, _env.TabPageEmail, "UiActive");
            }
            else
            {
            }

        }

        public void OnSUIDAdded()
        {
            object rewritten_ReadOnly = null;
            object NoPerson = null;
            object NoAsset = null;
            object GetLastSUIdx = null;
            object suindices = null;
            object agent = null;
            object Person = null;
            object helper = null;
            object responsibilty = null;
            object Anfrageart = null;
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            object responsibility = null; /* Undeclared in source */
            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v2 => { lcid = v2; })));

            rewritten_ReadOnly = true;
            NoPerson = true;
            NoAsset = true;

            //Zunächst überprüfen ob der Vorgang schreibgeschützt ist
            //First of all check whether the Case is write protected
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.hlobj, "IsReadOnly", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0))), (Int16)0)))
            {
                rewritten_ReadOnly = false;
            }

            //Status auf "In Bearbeitung" setzen
            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("IncidentStatusInProgress"));

            //Wenn Vorgang erweitert wird, wird die Zuständigkeit des Agenten ermittelt und gestezt.
            GetLastSUIdx = (Int16)0;
            suindices = _.VAL(_.CALL(this, _env.hlobj, "GetSvcUnitIndices", _.ARGS.ForceBrackets()));
            GetLastSUIdx = _.UBOUND(suindices);
            if (_.IF(_.GT(_.NullableNUM(GetLastSUIdx), (Int16)0)))
            {
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Val(_.ADD(GetLastSUIdx, (Int16)1)).Val((Int16)1)));
                helper = _.OBJ(_.CREATEOBJECT("helpline.hlcontrols.HLHelperPFA"));
                _env.Person = _.OBJ(_.CALL(this, helper, "GetPersonForAgent", _.ARGS.Val(_.CALL(this, _env.model, "GetClientContext")).Val(_.CLNG(agent))));
                if (_.IF(_.EQ(_.ISOBJECT(_env.Person), true)))
                {
                    responsibility = _.VAL(_.CALL(this, _env.Person, "GetValue", _.ARGS.Val("PersonGeneralTrumpf.Responsibility").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(responsibility), "ResponsibilityBSZDitzingen")))
                    {
                        _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.Responsibility").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ResponsibilityBSZDitzingen"));
                    }
                    else
                    {
                        _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.Responsibility").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ResponsibilityLocalIT"));
                    }
                }
            }

            //Zugriff auf Übersichts-Buttons regeln
            if (_.IF(_.EQ(rewritten_ReadOnly, false)))
            {
                _.SET(false, this, _env.ButtonShowOverView, "Disabled");
                _.SET(false, this, _env.ButtonEmailPreview, "Disabled");
                _.SET(false, this, _env.EditSubjectCase, "Disabled");
            }
            else
            {
                _.SET(true, this, _env.ButtonShowOverView, "Disabled");
                _.SET(true, this, _env.ButtonEmailPreview, "Disabled");
                _.SET(true, this, _env.EditSubjectCase, "Disabled");
            }
            //Abhängig von der Anfrageart werden Teile des Dialogs aktiviert oder deaktiviert
            Anfrageart = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeContact")))
            {
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }
            else
            {
                _.SET(true, this, _env.ComboIncidentStatus, "Disabled");
            }

            //Bei 2nd Level Dialog setzen der Benachrichtigung auf Email
            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("DefaultNotificationEmail"));

        }

        public void SearchAsset_AfterExecute()
        {
            //Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check Asset search status to set the caption of the button
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchAsset, "GetSearchState")), (Int16)3)))
            {
                _.SET("Reset", this, _env.SearchAsset, "Caption");
            }
            else
            {
                _.SET("Inventar", this, _env.SearchAsset, "Caption");
            }

        }

        public void SearchAsset_AfterReset()
        {
            object objO = null; /* Undeclared in source */
            object objT = null; /* Undeclared in source */
            objO = _.OBJ(_.CALL(this, _env.SearchAsset, "GetObject", _.ARGS.Val("product").Val(false)));
            objT = _.OBJ(_.CALL(this, _env.SearchAsset, "GetObject", _.ARGS.Val("product").Val(true)));

            _.CALL(this, objT, "SetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, objT, "SetValue", _.ARGS.Val("AssetGeneral.Hostname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, objT, "SetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));

            //Prüft ob Anfrager Objekt nicht vorhanden ist
            //Check wether the Caller object exist
            if (_.IF(_.OR(_.EQ(_.ISOBJECT(_env.hlcaller), false), _.EQ(_.NullableNUM(_.CALL(this, _env.hlcaller, "objID")), (Int16)0))))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            }

            //Status der Inventarsuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check Asset search status to set the caption of the button
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchAsset, "GetSearchState")), (Int16)3)))
            {
                _.SET("Reset", this, _env.SearchAsset, "Caption");
            }
            else
            {
                _.SET("Inventar", this, _env.SearchAsset, "Caption");
            }

        }

        public void SearchAsset_Click()
        {
            object rewritten_ReadOnly = null;
            object NoProduct = null;
            object hlProduct = null; /* Undeclared in source */
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            object varString = null; /* Undeclared in source */
            object varAType = null; /* Undeclared in source */
            rewritten_ReadOnly = true;
            NoProduct = true;

            //Wenn kein Inventar gefunden wurde, abbrechen
            //Cancel If no Asset was found
            if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, hlProduct, "GetType", _.ARGS.ForceBrackets())), "TEMPOBJECT")))
            {
                return;
            }

            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v3 => { lcid = v3; })));

            //Zunächst überprüfen ob der Vorgang schreibgeschützt ist
            //First of all check whether the Case is write protected
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.hlobj, "IsReadOnly", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0))), (Int16)0)))
            {
                rewritten_ReadOnly = false;
            }

            //Prüft ob ein Anfrager Objekt vorhanden ist und ob dieses auch angezeigt wird
            //Check wether the Caller object exist
            if (_.IF(_.AND(_.EQ(_.ISOBJECT(hlProduct), true), _.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditHostname, "Text")), ""))))
            {
                NoProduct = false;
            }

            if (_.IF(_.EQ(rewritten_ReadOnly, false)))
            {
                //Setzen des Inventars
                //Setting the asset
                varString = "";
                varAType = _.VAL(_.CALL(this, hlProduct, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.OR(_.OR(_.OR(_.EQ(_.NullableSTR(varAType), "DesktopComputer"), _.EQ(_.NullableSTR(varAType), "ServerComputer")), _.EQ(_.NullableSTR(varAType), "NotebookComputer")), _.EQ(_.NullableSTR(varAType), "Printer"))))
                {
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditHostname, "Text")), "")))
                    {
                        varString = _.VAL(_.CALL(this, _env.EditHostname, "Text"));
                    }
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditAssetModel, "Text")), "")))
                    {
                        varString = _.CONCAT(varString, " ", _.CALL(this, _env.EditAssetModel, "Text"));
                    }
                }
                else
                {
                    if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _env.EditAssetModel, "Text")), "")))
                    {
                        varString = _.VAL(_.CALL(this, _env.EditAssetModel, "Text"));
                    }
                    else
                    {
                        _.SET(" ", this, _env.EditAssetModel, "Text");
                    }
                }
                _.SET(_.VAL(varString), this, _env.EditAssetModel, "Text");
            }

        }

        public void SearchCaller_AfterExecute()
        {
            object tempmail = null;
            object strIncStatus = null;
            object CaseCallers = null;
            object Valid = null; /* Undeclared in source */
            object VIP = null; /* Undeclared in source */
            object sendmail = null; /* Undeclared in source */
            object strSubject = null; /* Undeclared in source */
            object strEmail = null; /* Undeclared in source */
            object CallerCount = null; /* Undeclared in source */
            object Caller = null; /* Undeclared in source */
            object CallerType = null; /* Undeclared in source */
            object mailadr = null; /* Undeclared in source */
            //Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check requester search status to set the caption of the button
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchCaller, "GetSearchState")), (Int16)3)))
            {
                _.SET("Reset", this, _env.SearchCaller, "Caption");
            }
            else
            {
                _.SET("Search", this, _env.SearchCaller, "Caption");
            }

            //VIP-Status des Anfragers abfragen und Imp Vorgang setzen
            Valid = _.VAL(_.CALL(this, _env.hlcaller, "HasContent", _.ARGS.Val("PersonGeneral.VIPLevel").Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableNUM(Valid), (Int16)1)))
            {
                VIP = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.VIPLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                //If VIP = VIPLevelNone Then hlObj.SetValue "IncidentAttribute.VIPStatus",0,0,0,"VIPStatusNone"
                if (_.IF(_.EQ(VIP, "VIPLevelVIP")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)1));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)142, (Int16)139, (Int16)254)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelITAdminDitzingen")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)2));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelITAdminTG")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)3));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelSAPKeyUserTUS")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)4));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET(_.VAL(_.RGB((Int16)205, (Int16)250, (Int16)255)), this, _env.Person, "BackColor");
                }
                else if (_.IF(_.EQ(VIP, "VIPLevelNon")))
                {
                    _.SET(false, this, _env.ComboVIPStatus, "Disabled");
                    _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)0));
                    _.SET(true, this, _env.ComboVIPStatus, "Disabled");
                    _.SET("", this, _env.Person, "BackColor");
                }
            }

            sendmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            tempmail = _.VAL(_.CALL(this, _env.EditEmailAddress, "text"));
            //Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
            //Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
            strIncStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strEmail = "";
            CallerCount = (Int16)0;
            CallerCount = _.VAL(_.CALL(this, _env.hlobj, "GetItemCount", _.ARGS.Val((Int16)0).Val((Int16)130)));

            if (_.IF(_.GT(_.NullableNUM(CallerCount), (Int16)0)))
            {
                CaseCallers = VBScriptConstants.Nothing;
                CaseCallers = _.VAL(_.CALL(this, _env.hlobj, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)130)));
                var enumerationContent = _.ENUMERABLE(CaseCallers).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent.MoveNext())
                        break;
                    Caller = enumerationContent.Current;
                    CallerType = _.VAL(_.CALL(this, Caller, "GetType"));
                    if (_.IF(_.EQ(_.NullableSTR(CallerType), "Employee")))
                    {
                        mailadr = "";
                        mailadr = _.VAL(_.CALL(this, Caller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                        if (_.IF(_.NOTEQ(_.NullableSTR(mailadr), "")))
                        {
                            strEmail = _.ADD(_.ADD(strEmail, mailadr), ";");
                        }
                    }
                }

            }
            else
            {
                strEmail = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            if (_.IF(_.GT(_.NullableNUM(_.INSTR(strEmail, tempmail)), (Int16)0)))
            {
            }
            else
            {
                strEmail = _.ADD(_.ADD(tempmail, ";"), strEmail);
            }

            if (_.IF(_.EQ(_.NullableSTR(strEmail), "")))
            {
                strEmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.EQ(_.NullableSTR(strEmail), "-")))
            {
                strEmail = "";
            }
            if (_.IF(_.EQ(_.NullableSTR(sendmail), "EmailCallerYes")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v4 => { strEmail = v4; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v5 => { strSubject = v5; }));
                _.SET(true, this, _env.TextBoxEmailTo, "Required");
                _.SET(true, this, _env.TextBoxEmailSubject, "Required");
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
            }
            else
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSearchName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSearchResult").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.SET(true, this, _env.GroupBoxEmail, "Disabled");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.TextBoxEmailSubject, "Required");
            }

        }

        public void SearchCaller_AfterReset()
        {
            object objO = null; /* Undeclared in source */
            object objT = null; /* Undeclared in source */
            objO = _.OBJ(_.CALL(this, _env.SearchCaller, "GetObject", _.ARGS.Val("caller").Val(false)));
            objT = _.OBJ(_.CALL(this, _env.SearchCaller, "GetObject", _.ARGS.Val("caller").Val(true)));

            _.CALL(this, objT, "SetValue", _.ARGS.Val("PersonGeneral.PersonSurname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, objT, "SetValue", _.ARGS.Val("PersonGeneral.PersonGivenName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, objT, "SetValue", _.ARGS.Val("PersonInformation.PersonOrganisation").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, objT, "SetValue", _.ARGS.Val("PersonInformation.PhoneNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));

            //Status der Anfragersuche prüfen, um die Bezeichnung des Suchbuttons zu setzen
            //Check requester search status to set the caption of the button
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.SearchCaller, "GetSearchState")), (Int16)3)))
            {
                _.SET("Reset", this, _env.SearchCaller, "Caption");
            }
            else
            {
                _.SET("", this, _env.EditSurname, "Text");
                _.SET("Search", this, _env.SearchCaller, "Caption");
            }

            //VIP-Status zurücksetzen
            _.CALL(this, _env.ComboVIPStatus, "SelectItem", _.ARGS.Val((Int16)0).Val((Int16)0));
            _.SET(_.VAL(_.RGB((Int16)248, (Int16)245, (Int16)240)), this, _env.Person, "BackColor");

        }

        public void SearchCaller_Click()
        {
            object rewritten_ReadOnly = null;
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            rewritten_ReadOnly = true;

            //Wenn keine Person gefunden wurde, abbrechen
            //Cancel If no person was found
            if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.hlcaller, "GetType", _.ARGS.ForceBrackets())), "TEMPOBJECT")))
            {
                return;
            }

            //Zunächst überprüfen ob der Vorgang schreibgeschützt ist
            //First of all check whether the Case is write protected
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.hlobj, "IsReadOnly", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0))), (Int16)0)))
            {
                rewritten_ReadOnly = false;
            }

            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v6 => { lcid = v6; })));

        }

        public void SetProblemText2Subject()
        {
            object varSubject = null; /* Undeclared in source */
            varSubject = _.VAL(_.LEFT(_.CALL(this, _env.EditProblem, "Text"), (Int16)100));
            if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.EditSubjectCase, "Text")), "")))
            {
                _.SET(_.REPLACE(varSubject, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " "), this, _env.EditSubjectCase, "Text");
            }

        }

        public void ComboIncidentStatus_SelectionChanged()
        {
            object tempmail = null;
            object strIncStatus = null;
            object CaseCallers = null;
            object strSubject = null; /* Undeclared in source */
            object strEmail = null; /* Undeclared in source */
            object CallerCount = null; /* Undeclared in source */
            object Caller = null; /* Undeclared in source */
            object CallerType = null; /* Undeclared in source */
            object mailadr = null; /* Undeclared in source */
            tempmail = _.VAL(_.CALL(this, _env.EditEmailAddress, "text"));
            strIncStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strEmail = "";
            CallerCount = (Int16)0;
            CallerCount = _.VAL(_.CALL(this, _env.hlobj, "GetItemCount", _.ARGS.Val((Int16)0).Val((Int16)130)));

            if (_.IF(_.GT(_.NullableNUM(CallerCount), (Int16)0)))
            {
                CaseCallers = VBScriptConstants.Nothing;
                CaseCallers = _.VAL(_.CALL(this, _env.hlobj, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)130)));
                var enumerationContent2 = _.ENUMERABLE(CaseCallers).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent2.MoveNext())
                        break;
                    Caller = enumerationContent2.Current;
                    CallerType = _.VAL(_.CALL(this, Caller, "GetType"));
                    if (_.IF(_.EQ(_.NullableSTR(CallerType), "Employee")))
                    {
                        mailadr = "";
                        mailadr = _.VAL(_.CALL(this, Caller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                        if (_.IF(_.NOTEQ(_.NullableSTR(mailadr), "")))
                        {
                            strEmail = _.ADD(_.ADD(strEmail, mailadr), ";");
                        }
                    }
                }
            }
            else
            {
                strEmail = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            if (_.IF(_.GT(_.NullableNUM(_.INSTR(strEmail, tempmail)), (Int16)0)))
            {
            }
            else
            {
                strEmail = _.ADD(_.ADD(tempmail, ";"), strEmail);
            }

            if (_.IF(_.EQ(_.NullableSTR(strEmail), "")))
            {
                strEmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.EQ(_.NullableSTR(strEmail), "-")))
            {
                strEmail = "";
            }
            if (_.IF(_.EQ(strIncStatus, "IncidentStatusSolved")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerYes"));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v7 => { strEmail = v7; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v8 => { strSubject = v8; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("SUINFO.PUBLISHED").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
                _.SET("Red", this, _env.LabelEmailBody, "TextColor");
                _.SET(true, this, _env.ComplexTextEmailBody, "Required");
                _.SET(true, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusClosed")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v9 => { strEmail = v9; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v10 => { strSubject = v10; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("SUINFO.PUBLISHED").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
                _.SET("Red", this, _env.LabelEmailBody, "TextColor");
                _.SET(true, this, _env.ComplexTextEmailBody, "Required");
                if (_.IF(_.EQ(_.NullableSTR(strEmail), "")))
                {
                    _.SET(false, this, _env.TextBoxEmailTo, "Required");
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerNo"));
                }
                else
                {
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerYes"));
                    _.SET(true, this, _env.TextBoxEmailTo, "Required");
                }
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusTimephased")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerYes"));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v11 => { strEmail = v11; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v12 => { strSubject = v12; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("SUINFO.PUBLISHED").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
                _.SET("Red", this, _env.LabelEmailBody, "TextColor");
                _.SET(true, this, _env.ComplexTextEmailBody, "Required");
                _.SET(true, this, _env.TextBoxEmailTo, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Disabled");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusWaitingforCustomer")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerYes"));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v13 => { strEmail = v13; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v14 => { strSubject = v14; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("SUINFO.PUBLISHED").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("1"));
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
                _.SET("Red", this, _env.LabelEmailBody, "TextColor");
                _.SET(true, this, _env.ComplexTextEmailBody, "Required");
                _.SET(true, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusWaitingforExtern")))
            {
                _.SET("Black", this, _env.LabelEmailBody, "TextColor");
                _.SET(false, this, _env.ComplexTextEmailBody, "Required");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusToProof")))
            {
                _.SET("Black", this, _env.LabelEmailBody, "TextColor");
                _.SET(false, this, _env.ComplexTextEmailBody, "Required");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusRouted")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerNo"));
                _.SET("Black", this, _env.LabelEmailBody, "TextColor");
                _.SET(false, this, _env.ComplexTextEmailBody, "Required");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusNew")))
            {
                _.SET("Black", this, _env.LabelEmailBody, "TextColor");
                _.SET(false, this, _env.ComplexTextEmailBody, "Required");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
            }
            else if (_.IF(_.EQ(strIncStatus, "IncidentStatusInProgress")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("EmailCallerNo"));
                _.SET("Black", this, _env.LabelEmailBody, "TextColor");
                _.SET(false, this, _env.ComplexTextEmailBody, "Required");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.EditResubmissionTime, "Required");
                _.SET(true, this, _env.EditResubmissionTime, "Disabled");
                _.CALL(this, _env.EditResubmissionTime, "DeleteContent");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
            }

        }

        public void ComboRequestType_SelectionChanged()
        {
            object Anfrageart = null;
            object Status = null;
            Anfrageart = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            Status = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            _.SET(false, this, _env.ComboProductionalRelevanz, "Disabled");

            if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeIncident")))
            {
                _.SET(true, this, _env.ComboImpact, "Disabled");
                _.SET(true, this, _env.ComboFunctionalRange, "Disabled");
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ImpactOne"));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("FunctionalRangePartFailure"));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ProductionalRelevanzAdministrativeProcess"));
            }
            else
            {
                _.SET(false, this, _env.ComboImpact, "Disabled");
                _.SET(false, this, _env.ComboFunctionalRange, "Disabled");
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ProductionalRelevanzSupportProcess"));
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeContact")))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                if (_.IF(_.NOTEQ(_.NullableSTR(Status), "IncidentStatusClosed")))
                {
                    _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                }
                else
                {
                    _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
                }
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }
            else
            {
                _.SET("", this, _env.EditProblem, "Text");
                _.SET(true, this, _env.CaseProblem, "Disabled");
                _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET("", this, _env.EditDiagnosis, "Text");
                _.SET(true, this, _env.CaseDiagnosis, "Disabled");
                _.SET(true, this, _env.KeywordTree, "Disabled");
                _.SET(true, this, _env.Attachment, "Disabled");
                _.SET(true, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboRequestType, "Disabled");
                _.SET(true, this, _env.ComboProductionalRelevanz, "Disabled");
                _.SET(true, this, _env.ComboIncidentStatus, "Disabled");
            }

        }

        public void OnSave()
        {
            object CheckOverView = null; /* Undeclared in source */
            object CheckSummaryHTML = null; /* Undeclared in source */
            //Priorität leeren, damit globale SLA´s auch runterstufen können
            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseClassificationAttribute.Priority").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("Priority5"));

            CheckOverView = "";
            CheckOverView = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.NOTEQ(_.NullableSTR(CheckOverView), "")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.Overview").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            }
            CheckSummaryHTML = "";
            CheckSummaryHTML = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.NOTEQ(_.NullableSTR(CheckSummaryHTML), "")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                //Button "Übersicht" entsperren
                _.SET(false, this, _env.ButtonShowOverView, "Disabled");
            }

        }

        public void TreeKeyword_ondatachange()
        {
            object isreserved = null;
            object agent = null;
            object agentid = null;
            object responsibility = null;
            object kw = null;
            object kwo = null;
            object cn = null; /* Undeclared in source */
            object rs_resp = null; /* Undeclared in source */
            object rs_kwkwo = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Bitte zuerst das Ticket reservieren.");
            }
            else
            {
                //Aktuellen Agent auslesen
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Datenbankverbindung zu helpline_replication
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Ditzingen oder TG auslesen
                rs_resp = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_resp = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select responsibility from AgentID_responsibility where agentid = ", _.CSTR(agent)))));
                responsibility = _.VAL(_.CALL(this, _.CALL(this, rs_resp, "fields", _.ARGS.Val("responsibility")), "value"));
                _.CALL(this, rs_resp, "close");

                //Keyword einlesen
                kw = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));
                if (_.IF(_.EQ(_.NullableNUM(responsibility), 112545)))
                {
                    //KeywordOrga Wert aus Vergleichstabelle einlesen
                    rs_kwkwo = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    rs_kwkwo = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select keywordorga from kw_kwo_mapping where keywordid = ", _.CSTR(kw)))));
                    while (_.IF(_.NOT(_.CALL(this, rs_kwkwo, "EOF"))))
                    {
                        kwo = _.VAL(_.CALL(this, _.CALL(this, rs_kwkwo, "fields", _.ARGS.Val("keywordorga")), "value"));
                        _.CALL(this, rs_kwkwo, "MoveNext");
                    }
                    if (_.IF(_.NOT(_.EQ(_.NullableSTR(kwo), ""))))
                    {
                        _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(kwo, v15 => { kwo = v15; }));
                        _.CALL(this, _env.TreeKeywordOrga, "SelectTreeItem", _.ARGS.Ref(kwo, v16 => { kwo = v16; }));
                    }
                    _.CALL(this, rs_kwkwo, "close");
                }
                else
                {
                    //Wert für die TG setzen
                    //Dim tg
                    //tg = HIER TG Value einlesen
                    //hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
                    //TreeKeywordOrga.SelectTreeItem tg
                }

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, cn, "close");
                cn = VBScriptConstants.Nothing;
            }

        }

        public void ComboLevel_SelectionChanged()
        {
            object level = null;
            //Bei Änderung des Supportlevels automatisch den Status auf "Weitergeleitet" setzen
            level = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.EscalationLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            if (_.IF(_.EQ(_.NullableSTR(level), "EscalationLevelLevel2")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("IncidentStatusRouted"));
            }
            if (_.IF(_.EQ(_.NullableSTR(level), "EscalationLevelLevel1")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("IncidentStatusRouted"));
            }

        }

        public void ButtonDiscovery_Click()
        {
            object Hostname = null;
            object wshshell = null;
            object oExec = null;
            object hlProduct = null; /* Undeclared in source */
            object Command1 = null; /* Undeclared in source */
            Hostname = _.VAL(_.CALL(this, hlProduct, "getvalue", _.ARGS.Val("AssetGeneral.Hostname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            wshshell = _.OBJ(_.CREATEOBJECT("Wscript.Shell"));
            Command1 = _.ADD(_.ADD("c:\\program files\\internet explorer\\iexplore.exe http://srv01inv1/discovery/Reports/List.aspx?q=", Hostname), "&flgDevice=1");
            oExec = _.OBJ(_.CALL(this, wshshell, "Exec", _.ARGS.Ref(Command1, v17 => { Command1 = v17; })));

        }

        public void b_template_save_Click()
        {
            object isreserved = null;
            object name = null;
            object agent = null;
            object teamID = null;
            object teamDisplayname = null;
            object agent_displayname = null;
            object result = null;
            object cn = null; /* Undeclared in source */
            object rs_team = null; /* Undeclared in source */
            object rs = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Please reserve the ticket first.");

            }
            else
            {

                //Templatenamen eingeben
                name = _.VAL(_.CALL(this, _, "INPUTBOX", _.ARGS.Val("Please type in a descriptive name for the template:").Val("templatename").Val("Maximum of 100 characters.")));

                //Bei Abbruch nichts unternehmen, sonst weiter im Script
                if (_.IF(_.EQ(name, false)))
                {
                }
                else
                {

                    //Agentid auslesen anhand des aktuellen Agenten
                    agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                    //Datenbankverbindung zu helpline_replication
                    cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                    //DB Verbindung öffnen
                    _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                    _.SET((Int16)10, this, cn, "ConnectionTimeout");
                    _.CALL(this, cn, "Open");

                    //Teamname auslesen
                    rs_team = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    rs_team = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select AgentTeam_ID,AgentTeam_Displayname,Agent_Displayname from IM_Agent_Supportteam where Agent_ID = ", _.CSTR(agent)))));
                    teamDisplayname = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("AgentTeam_Displayname")), "value"));
                    teamID = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("AgentTeam_ID")), "value"));
                    agent_displayname = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("Agent_Displayname")), "value"));
                    _.CALL(this, rs_team, "close");

                    //Abfrage ob Speicherung als persönliches oder als Teamtemplate gewünscht wird
                    result = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Button YES => personal template for: ", agent_displayname, _.CHR((Int16)10), _.CHR((Int16)13), _.CHR((Int16)13), "or", _.CHR((Int16)10), _.CHR((Int16)13), _.CHR((Int16)13), "Button NO => team template for: ''", teamDisplayname, "''")).Val((Int16)4).Val("personal template or team template?")));
                    if (_.IF(_.EQ(_.NullableNUM(result), (Int16)6)))
                    {
                        //Persönliches Insert auf Datenbank starten
                        rs = _.OBJ(_.CALL(this, cn, "execute", _.ARGS.Val(_.CONCAT("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('", _.CSTR(agent), "','", name, "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.EscalationLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CSTR(agent), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.Convenience").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.Rawtext").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "')"))));
                    }
                    else
                    {
                        //Team Insert auf Datenbank starten
                        rs = _.OBJ(_.CALL(this, cn, "execute", _.ARGS.Val(_.CONCAT("INSERT INTO templater (agentid, templatename,requesttype,descriptiontext,diagnosistext,solutiontext,keyword,keywordorga,escalationlevel,impact,functionalrange,productionalrelevance,emailcaller,incidentstatus,defaultnotification,editor,PCAssoziated,EmailBodyRawtext,EmailBodytext,EmailTo,EmailCC,EmailSubject) Values ('", _.CSTR(teamID), "','", name, "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.EscalationLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CSTR(agent), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.Convenience").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.Rawtext").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "','", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "')"))));

                    }
                    //Verbindung schließen
                    _.CALL(this, cn, "close");

                }
            }

        }

        public void b_template_load_Click()
        {
            object msg = null;
            object agent = null;
            object templateid = null;
            object tempmail = null;
            object CaseCallers = null;
            object Status = null;
            object cn = null; /* Undeclared in source */
            object rs = null; /* Undeclared in source */
            object strSubject = null; /* Undeclared in source */
            object varSubject = null; /* Undeclared in source */
            object strEmail = null; /* Undeclared in source */
            object CallerCount = null; /* Undeclared in source */
            object Caller = null; /* Undeclared in source */
            object CallerType = null; /* Undeclared in source */
            object mailadr = null; /* Undeclared in source */
            object sendmail = null; /* Undeclared in source */
            object Anfrageart = null; /* Undeclared in source */
            //Prüfen ob Template in der Checkbox ausgewählt wurde
            if (_.IF(_.OR(_.EQ(_.CALL(this, _env.cb_template_load, "GetCurSel"), (Int16)(-1)), _.EQ(_.NullableSTR(_.CALL(this, _env.l_templateID, "text")), ""))))
            {
                msg = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Please select a template from the list.", _.CHR((Int16)13), _.CHR((Int16)10), "If the list is empty, there is no template existing.")).Val(VBScriptConstants.vbOKOnly).Val("No data record available.")));
            }
            else
            {

                //Agentid auslesen anhand des aktuellen Agenten
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Angewählte ID aus Label auslesen
                templateid = _.VAL(_.CALL(this, _env.l_templateID, "Text"));

                //Datenbankverbindung zu helpline_replication
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Inhalte von agent_templates in das Recordset einlesen
                rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select * from templater where template_id = ", templateid))));

                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("Requesttype")), "value")));
                if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))), "")))
                {
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("descriptiontext")), "value")));
                }
                else
                {
                }
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("diagnosistext")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("solutiontext")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("keyword")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("keywordorga")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.EscalationLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EscalationLevel")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("Impact")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("FunctionalRange")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("ProductionalRelevance")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EmailCaller")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("IncidentStatus")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("DefaultNotification")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.Convenience").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("PCAssoziated")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EmailBodytext")), "value")));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EmailBodyRawtext")), "value")));
                //hlObj.SetValue "EmailSUAttribute.EmailTo",0,0,0,rs.fields("EmailTo").value
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EmailCC")), "value")));
                strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v18 => { strSubject = v18; }));
                if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))), "")))
                {
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("EmailSubject")), "value")));
                }

                //Subject Setzen
                varSubject = _.VAL(_.LEFT(_.CALL(this, _env.EditProblem, "Text"), (Int16)100));
                if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.EditSubjectCase, "Text")), "")))
                {
                    _.SET(_.REPLACE(varSubject, _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10)), " "), this, _env.EditSubjectCase, "Text");
                }

                //Übertrag der Caller in das An-Feld
                tempmail = _.VAL(_.CALL(this, _env.EditEmailAddress, "text"));
                strEmail = "";
                CallerCount = (Int16)0;
                CallerCount = _.VAL(_.CALL(this, _env.hlobj, "GetItemCount", _.ARGS.Val((Int16)0).Val((Int16)130)));

                if (_.IF(_.GT(_.NullableNUM(CallerCount), (Int16)0)))
                {
                    CaseCallers = VBScriptConstants.Nothing;
                    CaseCallers = _.VAL(_.CALL(this, _env.hlobj, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)130)));
                    var enumerationContent3 = _.ENUMERABLE(CaseCallers).GetEnumerator();
                    while (true)
                    {
                        if (!enumerationContent3.MoveNext())
                            break;
                        Caller = enumerationContent3.Current;
                        CallerType = _.VAL(_.CALL(this, Caller, "GetType"));
                        if (_.IF(_.EQ(_.NullableSTR(CallerType), "Employee")))
                        {
                            mailadr = "";
                            mailadr = _.VAL(_.CALL(this, Caller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                            if (_.IF(_.NOTEQ(_.NullableSTR(mailadr), "")))
                            {
                                strEmail = _.ADD(_.ADD(strEmail, mailadr), ";");
                            }
                        }
                    }
                }
                else
                {
                    strEmail = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                }

                if (_.IF(_.GT(_.NullableNUM(_.INSTR(strEmail, tempmail)), (Int16)0)))
                {
                }
                else
                {
                    strEmail = _.ADD(_.ADD(tempmail, ";"), strEmail);
                }

                if (_.IF(_.EQ(_.NullableSTR(strEmail), "")))
                {
                    strEmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                }
                if (_.IF(_.EQ(_.NullableSTR(strEmail), "-")))
                {
                    strEmail = "";
                }

                //Aktivieren der Felder je nach EmailCaller Wert
                sendmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.EQ(_.NullableSTR(sendmail), "EmailCallerYes")))
                {
                    _.SET(true, this, _env.TextBoxEmailTo, "Required");
                    _.SET(true, this, _env.TextBoxEmailSubject, "Required");
                    _.SET(false, this, _env.GroupBoxEmail, "Disabled");
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v19 => { strEmail = v19; }));
                }
                else
                {
                    _.SET(false, this, _env.TextBoxEmailTo, "Required");
                    _.SET(false, this, _env.TextBoxEmailSubject, "Required");
                    _.SET(true, this, _env.GroupBoxEmail, "Disabled");
                }

                //Aktivieren/Deaktivieren der Felder je nach gesetzter Anfrageart
                _.SET(false, this, _env.ComboProductionalRelevanz, "Disabled");
                if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeIncident")))
                {
                    _.SET(true, this, _env.ComboImpact, "Disabled");
                    _.SET(true, this, _env.ComboFunctionalRange, "Disabled");
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ImpactOne"));
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("FunctionalRangePartFailure"));
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ProductionalRelevanzAdministrativeProcess"));
                }
                else
                {
                    _.SET(false, this, _env.ComboImpact, "Disabled");
                    _.SET(false, this, _env.ComboFunctionalRange, "Disabled");
                    _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("ProductionalRelevanzSupportProcess"));
                }

                if (_.IF(_.NOTEQ(_.NullableSTR(Anfrageart), "RequestTypeContact")))
                {
                    _.SET(false, this, _env.CaseProblem, "Disabled");
                    Status = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.NOTEQ(_.NullableSTR(Status), "IncidentStatusClosed")))
                    {
                        _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                    }
                    else
                    {
                        _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
                    }
                    _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                    _.SET(false, this, _env.KeywordTree, "Disabled");
                    _.SET(false, this, _env.Attachment, "Disabled");
                    _.SET(false, this, _env.CaseAttributes, "Disabled");
                    _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
                }
                else
                {
                    _.SET("", this, _env.EditProblem, "Text");
                    _.SET(true, this, _env.CaseProblem, "Disabled");
                    _.SET(true, this, _env.ComboBoxEmailCaller, "Disabled");
                    _.SET("", this, _env.EditDiagnosis, "Text");
                    _.SET(true, this, _env.CaseDiagnosis, "Disabled");
                    _.SET(true, this, _env.KeywordTree, "Disabled");
                    _.SET(true, this, _env.Attachment, "Disabled");
                    _.SET(true, this, _env.CaseAttributes, "Disabled");
                    _.SET(false, this, _env.ComboRequestType, "Disabled");
                    _.SET(true, this, _env.ComboProductionalRelevanz, "Disabled");
                    _.SET(true, this, _env.ComboIncidentStatus, "Disabled");
                }

                //Recordset schließen
                _.CALL(this, rs, "close");
                rs = VBScriptConstants.Nothing;

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, cn, "close");
                cn = VBScriptConstants.Nothing;

            }

        }

        public void b_template_change_Click()
        {
            object agent = null;
            object templateid = null;
            object templatename = null;
            object editor = null;
            object agent_displayname = null;
            object msg2 = null;
            object name = null;
            object result = null;
            object cn = null; /* Undeclared in source */
            object rs = null; /* Undeclared in source */
            object rs_team = null; /* Undeclared in source */
            if (_.IF(_.OR(_.EQ(_.CALL(this, _env.cb_template_load, "GetCurSel"), (Int16)(-1)), _.EQ(_.NullableSTR(_.CALL(this, _env.l_templateID, "text")), ""))))
            {
                _.MSGBOX("Please select template from list first.");
            }
            else
            {

                //Agentid auslesen anhand des aktuellen Agenten
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Angewählte ID aus Label auslesen
                templateid = _.VAL(_.CALL(this, _env.l_templateID, "Text"));

                //Datenbankverbindung zu helpline_replication
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                //DB Verbindung öffnen
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Recordset anlegen und templatenamen auslesen
                rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select templatename,editor from templater where template_id = ", _.CSTR(templateid)))));
                templatename = _.VAL(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("templatename")), "value"));
                editor = _.VAL(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("editor")), "value"));

                //Agent Name auslesen
                rs_team = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_team = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select Agent_Displayname from IM_Agent_Supportteam where Agent_ID = ", _.CSTR(editor)))));
                agent_displayname = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("Agent_Displayname")), "value"));
                _.CALL(this, rs_team, "close");

                //Nur wenn Agent = Editor überschreiben, sonst Abbruch
                if (_.IF(_.NOTEQ(editor, _.CSTR(agent))))
                {
                    msg2 = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("You can only overwrite self-created templates.", _.CHR((Int16)10), _.CHR((Int16)13), "template: ", templateid, " was created by: ", agent_displayname, "")).Val(VBScriptConstants.vbOKOnly).Val("Overwrite is not allowed")));
                }
                else
                {
                    name = _.VAL(_.CALL(this, _, "INPUTBOX", _.ARGS.Val("Please type in a descriptive name: ").Val("overwrite template").Val(templatename)));
                    if (_.IF(_.EQ(name, false)))
                    {
                    }
                    else
                    {

                        //Abfrage ob Update erwünscht
                        result = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Möchten Sie das Template:  ''", templatename, "''  überschreiben?")).Val((Int16)4).Val("Template überschreiben?")));
                        if (_.IF(_.EQ(_.NullableNUM(result), (Int16)6)))
                        {

                            //Update auf Datenbank wird ausgeführt
                            rs = _.OBJ(_.CALL(this, cn, "execute", _.ARGS.Val(_.CONCAT("Update templater set templatename = '", name, "', Requesttype = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.RequestType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',descriptiontext = '", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "', diagnosistext = '", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "', solutiontext = '", _.REPLACE(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "'", "''"), "', keyword = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "', keywordorga = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "', EscalationLevel = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.EscalationLevel").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',Impact = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseClassificationAttribute.Impact").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',FunctionalRange = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.FunctionalRange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',ProductionalRelevance = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.ProductionalRelevanz").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailCaller = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',IncidentStatus = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',DefaultNotification = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',editor = '", _.CSTR(agent), "',PCAssoziated = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.Convenience").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailBodyRawtext = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.Rawtext").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailBodytext = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailTo = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailCC = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "',EmailSubject = '", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "' where template_id = ", _.CSTR(templateid)))));
                            rs = VBScriptConstants.Nothing;
                        }
                        else
                        {
                        }

                        //EndIF Überschreiben
                    }

                    //EndIf Agent = Editor
                }

                //Verbindung schließen
                _.CALL(this, cn, "close");

                //EndIf Wurde ein Checkbox-Wert zuvor angewählt
            }

        }

        public void cb_template_load_onfocus()
        {
            var errOn = _.GETERRORTRAPPINGTOKEN();
            object isreserved = null;
            object agent = null;
            object teamID = null;
            object teamDisplayname = null;
            object anzahl_agent_templates = null;
            object anzahl_team_templates = null;
            object cn = null; /* Undeclared in source */
            object rs_team = null; /* Undeclared in source */
            object rs = null; /* Undeclared in source */
            object rs2 = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Please reserve the ticket first.");
                _.SET(true, this, _env.EditSurname, "RequestFocus");
            }
            else
            {

                //Vorhandene Checkbox Werte entfernen
                _.CALL(this, _env.cb_template_load, "ResetContent");

                //Agentid auslesen anhand des aktuellen Agenten
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Datenbankverbindung zu helpline_replication
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Teamname auslesen
                rs_team = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_team = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select AgentTeam_ID,AgentTeam_Displayname from IM_Agent_Supportteam where Agent_ID = ", _.CSTR(agent)))));
                teamDisplayname = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("AgentTeam_Displayname")), "value"));
                teamID = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("AgentTeam_ID")), "value"));
                _.CALL(this, rs_team, "close");

                //Für Agent Templates ID bestimmen und selektierten Wert in Label schreiben
                anzahl_agent_templates = (Int16)0;
                rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select template_id,templatename from templater where agentid = ", _.CSTR(agent), " order by agentid, cast(Templatename as varchar(500))"))));
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, rs, "MoveFirst");
                });
                while (_.IF(() => _.IF(_.NOT(_.CALL(this, rs, "eof"))), errOn))
                {
                    _.HANDLEERROR(errOn, () => {
                        _.CALL(this, _env.cb_template_load, "AddItem", _.ARGS.Val(_.CALL(this, _.CALL(this, rs, "fields", _.ARGS.Val("templatename")), "value")));
                    });
                    _.HANDLEERROR(errOn, () => {
                        anzahl_agent_templates = _.ADD(anzahl_agent_templates, (Int16)1);
                    });
                    _.HANDLEERROR(errOn, () => {
                        _.CALL(this, rs, "MoveNext");
                    });
                }

                //Trennlinie zwischen Agent-Templates einfügen
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.cb_template_load, "AddItem", _.ARGS.Val("---------------------------------Team templates below---------------------------------"));
                });

                //Für Team Templates ID bestimmen und selektierten Wert in Label schreiben
                _.HANDLEERROR(errOn, () => {
                    anzahl_team_templates = (Int16)0;
                });
                _.HANDLEERROR(errOn, () => {
                    rs2 = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                });
                _.HANDLEERROR(errOn, () => {
                    rs2 = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select template_id,templatename from templater where agentid = ", _.CSTR(teamID), " order by agentid, cast(Templatename as varchar(500))"))));
                });
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, rs2, "MoveFirst");
                });
                while (_.IF(() => _.IF(_.NOT(_.CALL(this, rs2, "eof"))), errOn))
                {
                    _.HANDLEERROR(errOn, () => {
                        _.CALL(this, _env.cb_template_load, "AddItem", _.ARGS.Val(_.CALL(this, _.CALL(this, rs2, "fields", _.ARGS.Val("templatename")), "value")));
                    });
                    _.HANDLEERROR(errOn, () => {
                        anzahl_team_templates = _.ADD(anzahl_team_templates, (Int16)1);
                    });
                    _.HANDLEERROR(errOn, () => {
                        _.CALL(this, rs2, "MoveNext");
                    });
                }

                //Recordset schließen
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, rs, "close");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, rs2, "close");
                });

                //Datenbankverbindung zu helpline_replication schließen
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, cn, "close");
                });
                _.HANDLEERROR(errOn, () => {
                    cn = VBScriptConstants.Nothing;
                });

            }

            _.RELEASEERRORTRAPPINGTOKEN(errOn);
        }

        public void b_template_delete_Click()
        {
            object msg = null;
            object agent = null;
            object templateid = null;
            object editor = null;
            object agent_displayname = null;
            object msg2 = null;
            object result = null;
            object cn = null; /* Undeclared in source */
            object rs_editor = null; /* Undeclared in source */
            object rs_team = null; /* Undeclared in source */
            object rs = null; /* Undeclared in source */
            //Prüfen ob Template in der Checkbox ausgewählt wurde
            if (_.IF(_.OR(_.EQ(_.CALL(this, _env.cb_template_load, "GetCurSel"), (Int16)(-1)), _.EQ(_.NullableSTR(_.CALL(this, _env.l_templateID, "text")), ""))))
            {
                msg = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Please select a template from the list.", _.CHR((Int16)13), _.CHR((Int16)10), "If the list is empty, there is no template existing.")).Val(VBScriptConstants.vbOKOnly).Val("No data record available.")));

            }
            else
            {

                //Agentid auslesen anhand des aktuellen Agenten
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Angewählte ID aus Label auslesen
                templateid = _.VAL(_.CALL(this, _env.l_templateID, "Text"));

                //Datenbankverbindung zu helpline_replication
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Editor bestimmen
                rs_editor = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_editor = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select editor from templater where template_id = ", _.CSTR(templateid)))));
                editor = _.VAL(_.CALL(this, _.CALL(this, rs_editor, "fields", _.ARGS.Val("editor")), "value"));
                _.CALL(this, rs_editor, "close");

                //Agent Name auslesen
                rs_team = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_team = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select Agent_Displayname from IM_Agent_Supportteam where Agent_ID = ", _.CSTR(editor)))));
                agent_displayname = _.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("Agent_Displayname")), "value"));
                _.CALL(this, rs_team, "close");

                if (_.IF(_.NOTEQ(editor, _.CSTR(agent))))
                {
                    msg2 = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("You are only allowed to delete self-created tickets.", _.CHR((Int16)10), _.CHR((Int16)13), "Template ID: ", templateid, " was created by:", agent_displayname, "")).Val(VBScriptConstants.vbOKOnly).Val("Delete not allowed.")));
                }
                else
                {

                    //Abfrage ob Löschen erwünscht
                    result = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val("Do you really want to delete the template?").Val((Int16)4).Val("Delete template?")));
                    if (_.IF(_.EQ(_.NullableNUM(result), (Int16)6)))
                    {

                        //Zeile von agent_templates löschen
                        rs = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                        rs = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Delete from templater where template_id = ", _.CSTR(templateid)))));

                        //Auswahl der Checkbox zurücksetzen und ID auf Null
                        _.CALL(this, _env.cb_template_load, "ResetContent");
                        _.SET("", this, _env.l_templateID, "text");

                        //Recordset schließen
                        rs = VBScriptConstants.Nothing;
                    }
                    else
                    {
                    }

                    //End If Editor = Agent
                }

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, cn, "close");
                cn = VBScriptConstants.Nothing;

                //Vorhandene Checkbox Werte entfernen
                _.CALL(this, _env.cb_template_load, "ResetContent");
                _.SET("", this, _env.l_templateID, "Text");

            }

        }

        public void cb_template_load_SelectionEndOK()
        {
            var errOn2 = _.GETERRORTRAPPINGTOKEN();
            object agent = null;
            object team = null;
            object position = null;
            object teamID = null;
            object teamDisplayname = null;
            object anzahl_agent_templates = null;
            object cn = null; /* Undeclared in source */
            object rs_teamid = null; /* Undeclared in source */
            object rs_anzahl = null; /* Undeclared in source */
            object rs_agent = null; /* Undeclared in source */
            object i = null; /* Undeclared in source */
            object rs_team = null; /* Undeclared in source */
            //Agentid auslesen anhand des aktuellen Agenten
            agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

            //Angewählte Position bestimmen
            position = _.ADD(_.CALL(this, _env.cb_template_load, "GetCurSel"), (Int16)1);

            //Datenbankverbindung zu helpline_replication
            cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
            _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
            _.SET((Int16)10, this, cn, "ConnectionTimeout");
            _.CALL(this, cn, "Open");

            //Teamname auslesen
            rs_teamid = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
            rs_teamid = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select AgentTeam_ID,AgentTeam_Displayname from IM_Agent_Supportteam where Agent_ID = ", _.CSTR(agent)))));
            teamDisplayname = _.VAL(_.CALL(this, _.CALL(this, rs_teamid, "fields", _.ARGS.Val("AgentTeam_Displayname")), "value"));
            teamID = _.VAL(_.CALL(this, _.CALL(this, rs_teamid, "fields", _.ARGS.Val("AgentTeam_ID")), "value"));
            _.CALL(this, rs_teamid, "close");

            //Anzahl der Agenten-Templates bestimmen
            anzahl_agent_templates = (Int16)0;
            rs_anzahl = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
            rs_anzahl = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select template_id,templatename from templater where agentid = ", _.CSTR(agent)))));
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);
            _.HANDLEERROR(errOn2, () => {
                _.CALL(this, rs_anzahl, "MoveFirst");
            });
            while (_.IF(() => _.IF(_.NOT(_.CALL(this, rs_anzahl, "eof"))), errOn2))
            {
                _.HANDLEERROR(errOn2, () => {
                    anzahl_agent_templates = _.ADD(anzahl_agent_templates, (Int16)1);
                });
                _.HANDLEERROR(errOn2, () => {
                    _.CALL(this, rs_anzahl, "MoveNext");
                });
            }
            _.HANDLEERROR(errOn2, () => {
                _.CALL(this, rs_anzahl, "close");
            });

            if (_.IF(() => _.LTE(position, anzahl_agent_templates), errOn2))
            {
                //Select für Agententemplate ausführen
                _.HANDLEERROR(errOn2, () => {
                    rs_agent = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                });
                _.HANDLEERROR(errOn2, () => {
                    rs_agent = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select template_id from templater where agentid = '", _.CSTR(agent), "' order by agentid, cast(Templatename as varchar(500))"))));
                });
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);
                _.HANDLEERROR(errOn2, () => {
                    _.CALL(this, rs_agent, "MoveFirst");
                });
                object loopEnd = 0, loopStart = 0;
                var loopConstraintsInitialized = false;
                _.HANDLEERROR(errOn2, () => {
                    loopEnd = _.NUM(position);
                    loopStart = _.NUM((Int16)1);
                    if ((loopStart is DateTime) || (loopStart is Decimal))
                        i = loopStart;
                    loopStart = _.NUM((Int16)1, loopEnd);
                    loopConstraintsInitialized = true;
                });
                if (_.StrictLTE(loopStart, loopEnd))
                {
                    if (loopConstraintsInitialized)
                        i = loopStart;
                    while (true)
                    {
                        _.HANDLEERROR(errOn2, () => {
                            _.SET(_.VAL(_.CALL(this, _.CALL(this, rs_agent, "fields", _.ARGS.Val("template_id")), "value")), this, _env.l_templateID, "Text");
                        });
                        _.HANDLEERROR(errOn2, () => {
                            _.CALL(this, rs_agent, "MoveNext");
                        });
                        if (!loopConstraintsInitialized)
                            break;
                        var continueLoop = false;
                        _.HANDLEERROR(errOn2, () => {
                            i = _.ADD(i, (Int16)1);
                            continueLoop = _.StrictLTE(i, loopEnd);
                        });
                        if (!continueLoop)
                            break;
                    }
                }
                //Dataset schließen
                _.HANDLEERROR(errOn2, () => {
                    _.CALL(this, rs_agent, "close");
                });

            }
            else
            {

                //Prüfung, ob Trennlinie ausgewählt wurde.
                if (_.IF(() => _.EQ(_.CALL(this, _env.cb_template_load, "GetCurSel"), anzahl_agent_templates), errOn2))
                {
                    _.HANDLEERROR(errOn2, () => {
                        _.SET("", this, _env.l_templateID, "Text");
                    });
                    //cb_template_load.ResetContent

                }
                else
                {
                    //Select für Teamtemplate ausführen  - "Position -1" wegen Trennzeile zwischen Templatetypen
                    _.HANDLEERROR(errOn2, () => {
                        position = _.SUBT(_.SUBT(position, anzahl_agent_templates), (Int16)1);
                    });
                    _.HANDLEERROR(errOn2, () => {
                        rs_team = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    });
                    _.HANDLEERROR(errOn2, () => {
                        rs_team = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select template_id from templater where agentid = '", _.CSTR(teamID), "' order by agentid, cast(Templatename as varchar(500))"))));
                    });
                    _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);
                    _.HANDLEERROR(errOn2, () => {
                        _.CALL(this, rs_team, "MoveFirst");
                    });
                    object loopEnd2 = 0, loopStart2 = 0;
                    var loopConstraintsInitialized2 = false;
                    _.HANDLEERROR(errOn2, () => {
                        loopEnd2 = _.NUM(position);
                        loopStart2 = _.NUM((Int16)1);
                        if ((loopStart2 is DateTime) || (loopStart2 is Decimal))
                            i = loopStart2;
                        loopStart2 = _.NUM((Int16)1, loopEnd2);
                        loopConstraintsInitialized2 = true;
                    });
                    if (_.StrictLTE(loopStart2, loopEnd2))
                    {
                        if (loopConstraintsInitialized2)
                            i = loopStart2;
                        while (true)
                        {
                            _.HANDLEERROR(errOn2, () => {
                                _.SET(_.VAL(_.CALL(this, _.CALL(this, rs_team, "fields", _.ARGS.Val("template_id")), "value")), this, _env.l_templateID, "Text");
                            });
                            _.HANDLEERROR(errOn2, () => {
                                _.CALL(this, rs_team, "MoveNext");
                            });
                            if (!loopConstraintsInitialized2)
                                break;
                            var continueLoop2 = false;
                            _.HANDLEERROR(errOn2, () => {
                                i = _.ADD(i, (Int16)1);
                                continueLoop2 = _.StrictLTE(i, loopEnd2);
                            });
                            if (!continueLoop2)
                                break;
                        }
                    }
                    //Dataset schließen
                    _.HANDLEERROR(errOn2, () => {
                        _.CALL(this, rs_team, "close");
                    });

                }
            }

            //DB schließen
            _.HANDLEERROR(errOn2, () => {
                _.CALL(this, cn, "close");
            });

            _.RELEASEERRORTRAPPINGTOKEN(errOn2);
        }

        public void ButtonSCCMRemote_Click()
        {
            var errOn3 = _.GETERRORTRAPPINGTOKEN();
            object wshshell = null;
            object oExec = null;
            object OsType = null;
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            object objType = null; /* Undeclared in source */
            object hlProduct = null; /* Undeclared in source */
            object host = null; /* Undeclared in source */
            object Command1 = null; /* Undeclared in source */
            object RemoteTool = null; /* Undeclared in source */
            wshshell = _.OBJ(_.CREATEOBJECT("Wscript.Shell"));

            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v20 => { lcid = v20; })));

            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, _env.hlobj, "IsReadOnly", _.ARGS.Val("CASEINFO.REACTIONTIME").Val((Int16)0))), (Int16)0)))
            {

                objType = _.VAL(_.CALL(this, hlProduct, "GetType"));
                if (_.IF(_.OR(_.OR(_.EQ(_.NullableSTR(objType), "DesktopComputer"), _.EQ(_.NullableSTR(objType), "ServerComputer")), _.EQ(_.NullableSTR(objType), "NotebookComputer"))))
                {
                    //Auslesen des gewählten Computers
                    //Reading the selected computer
                    host = _.VAL(_.CALL(this, _env.EditHostname, "Text"));

                    if (_.IF(_.NOTEQ(_.NullableSTR(host), "")))
                    {
                        _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn3);
                        //Kommandozeile für den Aufruf von On Command Remote Master
                        //Command lin for calling On Command Remote Master
                        //Command1="""%programfiles%"\smsadmin\bin\i386\remote.exe 2 "" & host
                        _.HANDLEERROR(errOn3, () => {
                            OsType = _.VAL(_.CALL(this, _.GETOBJECT("winmgmts:root\\cimv2:Win32_Processor='cpu0'"), "AddressWidth"));
                        });
                        if (_.IF(() => _.EQ(_.NullableNUM(OsType), (Int16)32), errOn3))
                        {
                            //x86
                            _.HANDLEERROR(errOn3, () => {
                                Command1 = _.CONCAT("c:\\Program Files\\Microsoft Configuration Manager Console\\AdminUI\\bin\\i386\\rc.exe 1 ", host);
                            });
                        }
                        else
                        {
                            //x64
                            _.HANDLEERROR(errOn3, () => {
                                Command1 = _.CONCAT("c:\\Program Files (x86)\\Microsoft Configuration Manager Console\\AdminUI\\bin\\i386\\rc.exe 1 ", host);
                            });
                        }

                        _.HANDLEERROR(errOn3, () => {
                            RemoteTool = "SCCM Remote";
                        });

                        _.HANDLEERROR(errOn3, () => {
                            oExec = _.OBJ(_.CALL(this, wshshell, "Exec", _.ARGS.Ref(Command1, v21 => { Command1 = v21; })));
                        });
                        if (_.IF(() => _.EQ(_.CALL(this, _.ERR, "Number"), -2147024893), errOn3))
                        {
                            if (_.IF(() => _.EQ(_.NullableNUM(LangID), (Int16)7), errOn3))
                            {
                                _.HANDLEERROR(errOn3, () => {
                                    _.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Auf Ihrem Computer ist das Remote Tool ", RemoteTool, " nicht installiert.", VBScriptConstants.vbLf, "Bitte wenden Sie sich an Ihren Administrator.")).Val(VBScriptConstants.vbExclamation).Val("helpLine - ClassicDesk"));
                                });
                            }
                            else
                            {
                                _.HANDLEERROR(errOn3, () => {
                                    _.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("The remote tool ", RemoteTool, " is not installed on your computer.", VBScriptConstants.vbLf, "Please consult your administrator.")).Val(VBScriptConstants.vbExclamation).Val("helpLine - ClassicDesk"));
                                });
                            }
                        }
                    }
                }
                else
                {
                    if (_.IF(() => _.EQ(_.NullableNUM(LangID), (Int16)7), errOn3))
                    {
                        _.HANDLEERROR(errOn3, () => {
                            _.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("Es wurde kein Computer als Inventar ausgewählt.", VBScriptConstants.vbLf, "Bitte wählen Sie einen Computer für den Vorgang aus.")).Val(VBScriptConstants.vbExclamation).Val("helpLine - ClassicDesk"));
                        });
                    }
                    else
                    {
                        _.HANDLEERROR(errOn3, () => {
                            _.CALL(this, _, "MSGBOX", _.ARGS.Val(_.CONCAT("No computer has been selected.", VBScriptConstants.vbLf, "Please select a computer for this Case.")).Val(VBScriptConstants.vbExclamation).Val("helpLine - ClassicDesk"));
                        });
                    }
                }
            }

            _.RELEASEERRORTRAPPINGTOKEN(errOn3);
        }

        public void ButtonShowOverView_Click()
        {
            object lcid = null; /* Undeclared in source */
            object LangID = null; /* Undeclared in source */
            object CaseOwner = null; /* Undeclared in source */
            object agent = null; /* Undeclared in source */
            object Problemtitle = null; /* Undeclared in source */
            object Diagnosistitle = null; /* Undeclared in source */
            object Solutiontitle = null; /* Undeclared in source */
            object DescrText = null; /* Undeclared in source */
            object ProblemAll = null; /* Undeclared in source */
            object actStatus = null; /* Undeclared in source */
            object SolText = null; /* Undeclared in source */
            object SolutionAll = null; /* Undeclared in source */
            object SUIdx = null; /* Undeclared in source */
            object SUDiagnosisIntern = null; /* Undeclared in source */
            object SUDiagnosis = null; /* Undeclared in source */
            object i = null; /* Undeclared in source */
            object SUDiagnosisExtern = null; /* Undeclared in source */
            object SUDiagnosisExt = null; /* Undeclared in source */
            object SUActivity = null; /* Undeclared in source */
            object SURegTime = null; /* Undeclared in source */
            object DiagnosisAll = null; /* Undeclared in source */
            //Ermitteln der Locale ID für die Sprachauswahl
            //Selecting the Locale ID for the desired language
            lcid = _.VAL(_.CALL(this, _env.hlSession, "GetLocaleID"));
            LangID = _.VAL(_.CALL(this, _env.hlSession, "LangIDFromLCID", _.ARGS.Ref(lcid, v22 => { lcid = v22; })));

            CaseOwner = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("HLOBJECTINFO.OWNER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            agent = "";
            if (_.IF(_.EQ(_.NullableNUM(LangID), (Int16)7)))
            {
                Problemtitle = _.CONCAT("<b>====== Problembeschreibung ======", " [von Agent : ", CaseOwner, "]</b>", VBScriptConstants.vbNewLine);
                Diagnosistitle = _.CONCAT("<b>====== Kommunikation ======</b>", VBScriptConstants.vbNewLine);
                Solutiontitle = _.CONCAT("<b>====== Lösungsbeschreibung ======", " [von Agent : ", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "]</b>", VBScriptConstants.vbNewLine);
            }
            else
            {
                Problemtitle = _.CONCAT("<b>====== Problemdescription ======", " [by Agent : ", CaseOwner, "]</b>", VBScriptConstants.vbNewLine);
                Diagnosistitle = _.CONCAT("<b>====== Diagnosisactivities ======</b>", VBScriptConstants.vbNewLine);
                Solutiontitle = _.CONCAT("<b>====== Final solution ======", " [by Agent : ", _.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), "]</b>", VBScriptConstants.vbNewLine);
            }
            //VG-Beschreibung
            DescrText = "";
            DescrText = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)1).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(DescrText), "")))
            {
                DescrText = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.NOTEQ(_.NullableSTR(DescrText), "")))
            {
                DescrText = _.REPLACE(DescrText, VBScriptConstants.vbCrLf, "<br>");
                ProblemAll = _.CONCAT(Problemtitle, DescrText, VBScriptConstants.vbNewLine);
            }
            //VG-Lösung
            //nur bei Status "Geschlossen" aus der aktuellen SU den Text holen
            actStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            SolText = "";
            if (_.IF(_.EQ(_.NullableSTR(actStatus), "IncidentStatusClosed")))
            {
                SolText = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.EQ(_.NullableSTR(SolText), "")))
            {
                SolText = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseSolution.SolutionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.NOTEQ(_.NullableSTR(SolText), "")))
            {
                SolText = _.REPLACE(SolText, VBScriptConstants.vbCrLf, "<br>");
                SolutionAll = _.CONCAT(Solutiontitle, SolText);
            }

            SUIdx = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.INDEX").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.GT(_.NullableNUM(SUIdx), (Int16)0)))
            {
                //Pro SU prüfen, ob Tätigkeitsbeschreibung eingetragen ist
                var loopEnd3 = _.NUM(SUIdx);
                var loopStart3 = _.NUM((Int16)1, loopEnd3);
                if (_.StrictLTE(loopStart3, loopEnd3))
                {
                    for (i = loopStart3; _.StrictLTE(i, loopEnd3); i = _.ADD(i, (Int16)1))
                    {
                        SUDiagnosisIntern = "<b> --- intern --- </b>";
                        SUDiagnosis = "";
                        SUDiagnosis = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Ref(i, v23 => { i = v23; }).Val((Int16)0)));
                        //SUDiagnosis = Replace(SUDiagnosis, Chr(13) & Chr(10), " ")
                        SUDiagnosis = _.REPLACE(SUDiagnosis, VBScriptConstants.vbCrLf, "<br>");
                        SUDiagnosisExtern = "<b> --- extern --- </b>";
                        SUDiagnosisExt = "";
                        SUDiagnosisExt = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Ref(i, v24 => { i = v24; }).Val((Int16)0)));
                        if (_.IF(_.NOTEQ(_.NullableSTR(SUDiagnosis), "")))
                        {
                            SUActivity = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentSUAttribute.IncidentOperation").Ref(LangID, v25 => { LangID = v25; }).Val((Int16)0).Ref(i, v26 => { i = v26; }).Val((Int16)0)));
                            SURegTime = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(i, v27 => { i = v27; }).Val((Int16)0)));
                            agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(i, v28 => { i = v28; }).Val((Int16)0)));
                            DiagnosisAll = _.CONCAT(DiagnosisAll, SUDiagnosisIntern, VBScriptConstants.vbNewLine, "<b>", i, ". SU (", agent, ") -> ", SUActivity, " [", SURegTime, "]:", "</b>", VBScriptConstants.vbNewLine, SUDiagnosis, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                        }
                        if (_.IF(_.NOTEQ(_.NullableSTR(SUDiagnosisExt), "")))
                        {
                            //SUDiagnosisExt = Replace(SUDiagnosisExt, vbCrLf, "<br>")
                            SUActivity = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentSUAttribute.IncidentOperation").Ref(LangID, v29 => { LangID = v29; }).Val((Int16)0).Ref(i, v30 => { i = v30; }).Val((Int16)0)));
                            SURegTime = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.REGISTRATIONTIME").Val((Int16)0).Val((Int16)0).Ref(i, v31 => { i = v31; }).Val((Int16)0)));
                            agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Ref(i, v32 => { i = v32; }).Val((Int16)0)));
                            DiagnosisAll = _.CONCAT(DiagnosisAll, SUDiagnosisExtern, VBScriptConstants.vbNewLine, "<b>", i, ". SU (", agent, ") -> ", SUActivity, " [", SURegTime, "]:", "</b>", VBScriptConstants.vbNewLine, SUDiagnosisExt, VBScriptConstants.vbNewLine, _.STRING((Int16)80, "-"), VBScriptConstants.vbNewLine);
                        }
                    }
                }
            }
            if (_.IF(_.NOTEQ(_.NullableSTR(DiagnosisAll), "")))
            {
                DiagnosisAll = _.CONCAT(Diagnosistitle, DiagnosisAll);
            }
            ProblemAll = _.CONCAT(ProblemAll, DiagnosisAll, SolutionAll);
            //hlObj.SetValue "CaseGeneral.Overview",0,0,0,ProblemAll
            ProblemAll = _.REPLACE(ProblemAll, VBScriptConstants.vbCrLf, "<br>");
            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(ProblemAll, v33 => { ProblemAll = v33; }));

            //Button nach 1. Klick sperren
            //ButtonShowOverView.Disabled = True

        }

        public void ComboBoxEmailCaller_SelectionChanged()
        {
            object tempmail = null;
            object strIncStatus = null;
            object CaseCallers = null;
            object sendmail = null; /* Undeclared in source */
            object strSubject = null; /* Undeclared in source */
            object strEmail = null; /* Undeclared in source */
            object CallerCount = null; /* Undeclared in source */
            object Caller = null; /* Undeclared in source */
            object CallerType = null; /* Undeclared in source */
            object mailadr = null; /* Undeclared in source */
            sendmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailCaller").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            tempmail = _.VAL(_.CALL(this, _env.EditEmailAddress, "text"));
            //Rote Titel-Beschriftung des Lösungstextfeldes bei Inc.-Status Gelöst/Geschlosssen.
            //Redcoloured title of the solutiontext-frame if Inc.-status Solved or Closed.
            strIncStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strSubject = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            strEmail = "";
            CallerCount = (Int16)0;
            CallerCount = _.VAL(_.CALL(this, _env.hlobj, "GetItemCount", _.ARGS.Val((Int16)0).Val((Int16)130)));

            if (_.IF(_.GT(_.NullableNUM(CallerCount), (Int16)0)))
            {
                CaseCallers = VBScriptConstants.Nothing;
                CaseCallers = _.VAL(_.CALL(this, _env.hlobj, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)130)));
                var enumerationContent4 = _.ENUMERABLE(CaseCallers).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent4.MoveNext())
                        break;
                    Caller = enumerationContent4.Current;
                    CallerType = _.VAL(_.CALL(this, Caller, "GetType"));
                    if (_.IF(_.EQ(_.NullableSTR(CallerType), "Employee")))
                    {
                        mailadr = "";
                        mailadr = _.VAL(_.CALL(this, Caller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                        if (_.IF(_.NOTEQ(_.NullableSTR(mailadr), "")))
                        {
                            strEmail = _.ADD(_.ADD(strEmail, mailadr), ";");
                        }
                    }
                }

            }
            else
            {
                strEmail = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            if (_.IF(_.GT(_.NullableNUM(_.INSTR(strEmail, tempmail)), (Int16)0)))
            {
            }
            else
            {
                strEmail = _.ADD(_.ADD(tempmail, ";"), strEmail);
            }

            if (_.IF(_.EQ(_.NullableSTR(strEmail), "")))
            {
                strEmail = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            if (_.IF(_.EQ(_.NullableSTR(strEmail), "-")))
            {
                strEmail = "";
            }
            if (_.IF(_.EQ(_.NullableSTR(sendmail), "EmailCallerYes")))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v34 => { strEmail = v34; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strSubject, v35 => { strSubject = v35; }));
                _.SET(true, this, _env.TextBoxEmailTo, "Required");
                _.SET(true, this, _env.TextBoxEmailSubject, "Required");
                _.SET(false, this, _env.GroupBoxEmail, "Disabled");
            }
            else
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSearchName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSearchResult").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailCC").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
                _.SET(true, this, _env.GroupBoxEmail, "Disabled");
                _.SET(false, this, _env.TextBoxEmailTo, "Required");
                _.SET(false, this, _env.TextBoxEmailSubject, "Required");
            }

        }

        public void ButtonSearchMail_Click()
        {
            object ConString = null;
            object name = null; /* Undeclared in source */
            object cn2 = null; /* Undeclared in source */
            object rs2 = null; /* Undeclared in source */
            object Data = null; /* Undeclared in source */
            object i = null; /* Undeclared in source */
            //EMail-Adressen leeren
            _.SET("", this, _env.ComboBoxEmailSearchResult, "Text");
            _.CALL(this, _env.ComboBoxEmailSearchResult, "ResetContent");
            //Name als Suchparameter für Email-Adressen abfragen
            name = _.VAL(_.CALL(this, _env.TextBoxEmailSearchName, "Text"));

            //ConString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm4t"
            ConString = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm1";

            //Dim i
            //i = 0

            if (_.IF(_.NOTEQ(_.NullableSTR(name), "")))
            {
                //------------------------------------------------------------------------------------------------
                //Ermitteln der Email-Adressen auf Bases des eingegebenen Namens
                cn2 = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));

                //Verbindung öffnen
                _.SET(_.VAL(ConString), this, cn2, "ConnectionString");
                _.SET((Int16)10, this, cn2, "ConnectionTimeout");
                _.CALL(this, cn2, "Open");

                //SELECT absetzen
                rs2 = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs2 = _.OBJ(_.CALL(this, cn2, "Execute", _.ARGS.Val(_.CONCAT("select email from _EMails where email LIKE '%", name, "%' ORDER BY email"))));

                //Daten einlesen
                Data = "";
                while (_.IF(_.NOT(_.CALL(this, rs2, "eof"))))
                {
                    //In Variable schreiben
                    i = _.ADD(i, (Int16)1);
                    _.CALL(this, _env.ComboBoxEmailSearchResult, "AddItem", _.ARGS.Val(_.CALL(this, _.CALL(this, rs2, "fields", _.ARGS.Val((Int16)0)), "value")));
                    if (_.IF(_.EQ(_.NullableNUM(i), (Int16)1)))
                    {
                        _.SET(_.VAL(_.CALL(this, _.CALL(this, rs2, "fields", _.ARGS.Val((Int16)0)), "value")), this, _env.ComboBoxEmailSearchResult, "Text");
                    }
                    _.CALL(this, rs2, "movenext");
                }
                //Verbindung schließen
                _.CALL(this, rs2, "close");
                _.CALL(this, cn2, "close");

            }

        }

        public void ButtonTo_Click()
        {
            object email = null; /* Undeclared in source */
            object Recipient = null; /* Undeclared in source */
            object fullemailstring = null; /* Undeclared in source */
            object pos = null; /* Undeclared in source */
            object emailstring = null; /* Undeclared in source */
            email = _.VAL(_.CALL(this, _env.ComboBoxEmailSearchResult, "Text"));
            Recipient = _.VAL(_.CALL(this, _env.TextBoxEmailTo, "Text"));
            if (_.IF(_.EQ(_.NullableSTR(email), "")))
            {
                _.MSGBOX("Bitte eine Email-Adresse auswählen!");
            }
            else
            {
                fullemailstring = _.VAL(_.LEN(email));
                pos = _.VAL(_.INSTR((Int16)1, email, ":", (Int16)1));
                emailstring = _.SUBT(_.CLNG(fullemailstring), _.CLNG(pos));
                email = _.VAL(_.RIGHT(email, _.CLNG(emailstring)));
                if (_.IF(_.EQ(_.NullableSTR(Recipient), "")))
                {
                    Recipient = _.VAL(email);
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(_.RIGHT(Recipient, (Int16)1)), ";")))
                    {
                        Recipient = _.ADD(Recipient, email);
                    }
                    else
                    {
                        Recipient = _.ADD(_.ADD(Recipient, ";"), email);
                    }
                }
                _.SET(_.VAL(Recipient), this, _env.TextBoxEmailTo, "Text");
            }

        }

        public void ButtonCC_Click()
        {
            object email = null; /* Undeclared in source */
            object RecipientCC = null; /* Undeclared in source */
            object fullemailstring = null; /* Undeclared in source */
            object pos = null; /* Undeclared in source */
            object emailstring = null; /* Undeclared in source */
            email = _.VAL(_.CALL(this, _env.ComboBoxEmailSearchResult, "Text"));
            RecipientCC = _.VAL(_.CALL(this, _env.TextBoxEmailCC, "Text"));
            if (_.IF(_.EQ(_.NullableSTR(email), "")))
            {
                _.MSGBOX("Bitte eine Email-Adresse auswählen!");
            }
            else
            {
                fullemailstring = _.VAL(_.LEN(email));
                pos = _.VAL(_.INSTR((Int16)1, email, ":", (Int16)1));
                emailstring = _.SUBT(_.CLNG(fullemailstring), _.CLNG(pos));
                email = _.VAL(_.RIGHT(email, _.CLNG(emailstring)));
                if (_.IF(_.EQ(_.NullableSTR(RecipientCC), "")))
                {
                    RecipientCC = _.VAL(email);
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(_.RIGHT(RecipientCC, (Int16)1)), ";")))
                    {
                        RecipientCC = _.ADD(RecipientCC, email);
                    }
                    else
                    {
                        RecipientCC = _.ADD(_.ADD(RecipientCC, ";"), email);
                    }
                }
                _.SET(_.VAL(RecipientCC), this, _env.TextBoxEmailCC, "Text");
            }

        }

        public void ButtonSetAgent1_Click()
        {
            object isreserved = null;
            object agent = null;
            object agentid = null;
            object internalname = null;
            object cn = null; /* Undeclared in source */
            object rs_kwo = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Bitte zuerst das Ticket reservieren.");
            }
            else
            {

                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Datenbankverbindung zu helpline_data
                cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn, "ConnectionString");
                _.SET((Int16)10, this, cn, "ConnectionTimeout");
                _.CALL(this, cn, "Open");

                //Teamname auslesen
                rs_kwo = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_kwo = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Select name,internalname from vw_agent_to_first_keywordorga where agentid = ", _.CSTR(agent)))));
                internalname = _.VAL(_.CALL(this, _.CALL(this, rs_kwo, "fields", _.ARGS.Val("internalname")), "value"));

                //Wert in Schlagwort schreiben
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(internalname, v36 => { internalname = v36; }));
                _.CALL(this, _env.TreeKeywordOrga, "SelectTreeItem", _.ARGS.Ref(internalname, v37 => { internalname = v37; }));

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, rs_kwo, "close");
                _.CALL(this, cn, "close");
                cn = VBScriptConstants.Nothing;

            }

        }

        public void ButtonSetKW_Click()
        {
            object isreserved = null;
            object agent = null;
            object keywordid = null;
            object responsibility = null;
            object kw = null;
            object kwo = null;
            object cn1 = null; /* Undeclared in source */
            object rs_kw = null; /* Undeclared in source */
            object rs_resp = null; /* Undeclared in source */
            object rs_kwkwo = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Bitte zuerst das Ticket reservieren.");
            }
            else
            {
                //Aktuellen Agent auslesen
                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Datenbankverbindung zu helpline_replication
                cn1 = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn1, "ConnectionString");
                _.SET((Int16)10, this, cn1, "ConnectionTimeout");
                _.CALL(this, cn1, "Open");

                //Teamname auslesen
                rs_kw = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_kw = _.OBJ(_.CALL(this, cn1, "Execute", _.ARGS.Val(_.CONCAT("Select keywordid from vw_Agent_Emplkeyword where agentid = ", _.CSTR(agent)))));
                keywordid = _.VAL(_.CALL(this, _.CALL(this, rs_kw, "fields", _.ARGS.Val("keywordid")), "value"));
                _.CALL(this, rs_kw, "close");

                //Wert in Schlagwort schreiben
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(keywordid, v38 => { keywordid = v38; }));
                _.CALL(this, _env.TreeKeyword, "SelectTreeItem", _.ARGS.Ref(keywordid, v39 => { keywordid = v39; }));
                _.CALL(this, _env.TreeKeyword, "ExpandTreeItem", _.ARGS.Ref(keywordid, v40 => { keywordid = v40; }));

                //Responsibility - Ditzingen oder TG - einlesen
                rs_resp = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                rs_resp = _.OBJ(_.CALL(this, cn1, "Execute", _.ARGS.Val(_.CONCAT("Select responsibility from AgentID_responsibility where agentid = ", _.CSTR(agent)))));
                responsibility = _.VAL(_.CALL(this, _.CALL(this, rs_resp, "fields", _.ARGS.Val("responsibility")), "value"));
                _.CALL(this, rs_resp, "close");

                //Keyword einlesen
                kw = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));
                if (_.IF(_.EQ(_.NullableNUM(responsibility), 112545)))
                {
                    //KeywordOrga Wert aus Vergleichstabelle einlesen
                    rs_kwkwo = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    rs_kwkwo = _.OBJ(_.CALL(this, cn1, "Execute", _.ARGS.Val(_.CONCAT("Select keywordorga from kw_kwo_mapping where keywordid = ", _.CSTR(kw)))));
                    while (_.IF(_.NOT(_.CALL(this, rs_kwkwo, "EOF"))))
                    {
                        kwo = _.VAL(_.CALL(this, _.CALL(this, rs_kwkwo, "fields", _.ARGS.Val("keywordorga")), "value"));
                        _.CALL(this, rs_kwkwo, "MoveNext");
                    }
                    if (_.IF(_.NOT(_.EQ(_.NullableSTR(kwo), ""))))
                    {
                        _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(kwo, v41 => { kwo = v41; }));
                        _.CALL(this, _env.TreeKeywordOrga, "SelectTreeItem", _.ARGS.Ref(kwo, v42 => { kwo = v42; }));
                    }
                    _.CALL(this, rs_kwkwo, "close");
                }
                else
                {
                    //Wert für die TG setzen
                    //Dim tg
                    //tg = HIER TG Value einlesen
                    //hlObj.SetValue "Keywords.KeywordOrga",0,0,0,tg
                    //TreeKeywordOrga.SelectTreeItem tg
                }

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, cn1, "close");
                cn1 = VBScriptConstants.Nothing;
            }

        }

        public void ButtonResetTo_Click()
        {
            object CaseCallers = null;
            object tempmail = null;
            object CallerCount = null; /* Undeclared in source */
            object Caller = null; /* Undeclared in source */
            object CallerType = null; /* Undeclared in source */
            object mailadr = null; /* Undeclared in source */
            object strEmail = null; /* Undeclared in source */
            CallerCount = (Int16)0;
            CallerCount = _.VAL(_.CALL(this, _env.hlobj, "GetItemCount", _.ARGS.Val((Int16)0).Val((Int16)130)));
            if (_.IF(_.GT(_.NullableNUM(CallerCount), (Int16)0)))
            {
                CaseCallers = VBScriptConstants.Nothing;
                CaseCallers = _.VAL(_.CALL(this, _env.hlobj, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)130)));
                var enumerationContent5 = _.ENUMERABLE(CaseCallers).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent5.MoveNext())
                        break;
                    Caller = enumerationContent5.Current;
                    CallerType = _.VAL(_.CALL(this, Caller, "GetType"));
                    if (_.IF(_.EQ(_.NullableSTR(CallerType), "Employee")))
                    {
                        mailadr = "";
                        mailadr = _.VAL(_.CALL(this, Caller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                        if (_.IF(_.NOTEQ(_.NullableSTR(mailadr), "")))
                        {
                            strEmail = _.ADD(_.ADD(strEmail, mailadr), ";");
                        }
                    }
                }
            }
            else
            {
                strEmail = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            tempmail = _.VAL(_.CALL(this, _env.EditEmailAddress, "text"));
            if (_.IF(_.GT(_.NullableNUM(_.INSTR(strEmail, tempmail)), (Int16)0)))
            {
            }
            else
            {
                strEmail = _.ADD(_.ADD(tempmail, ";"), strEmail);
            }

            _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strEmail, v43 => { strEmail = v43; }));

        }

        public void ButtonEmailPreview_Click()
        {
            object ForReading = null;
            object ForWriting = null;
            object ForAppending = null;
            object OriginDescr = null;
            object Status = null; /* Undeclared in source */
            object HLinkToCase = null; /* Undeclared in source */
            object HTicketID = null; /* Undeclared in source */
            object SubjectCase = null; /* Undeclared in source */
            object LanguageDE = null; /* Undeclared in source */
            object MailTo = null; /* Undeclared in source */
            object z = null; /* Undeclared in source */
            object CounterEmpf = null; /* Undeclared in source */
            object surname = null; /* Undeclared in source */
            object letteraddress = null; /* Undeclared in source */
            object language = null; /* Undeclared in source */
            object editor = null; /* Undeclared in source */
            object MailBody = null; /* Undeclared in source */
            object TTicketID = null; /* Undeclared in source */
            object TStatus = null; /* Undeclared in source */
            object HStatus = null; /* Undeclared in source */
            object LastSUIdx = null; /* Undeclared in source */
            object TEditor = null; /* Undeclared in source */
            object TSubject = null; /* Undeclared in source */
            object Anrede = null; /* Undeclared in source */
            object TSolution = null; /* Undeclared in source */
            object TBeschr = null; /* Undeclared in source */
            object TComplimentary = null; /* Undeclared in source */
            object TSignature = null; /* Undeclared in source */
            object TNoticeTop = null; /* Undeclared in source */
            object Creationdate = null; /* Undeclared in source */
            object Datum = null; /* Undeclared in source */
            object subject = null; /* Undeclared in source */
            object TIntroduction = null; /* Undeclared in source */
            object fso = null; /* Undeclared in source */
            object f = null; /* Undeclared in source */
            object BodyText = null; /* Undeclared in source */
            object TNoticeBottom = null; /* Undeclared in source */
            object DiagnText = null; /* Undeclared in source */
            object TDiagnosis = null; /* Undeclared in source */
            object TResubTime = null; /* Undeclared in source */
            object ResubmissionTime = null; /* Undeclared in source */
            object ResubDatum = null; /* Undeclared in source */
            Status = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            HLinkToCase = "http://srv01itsm2/helpLinePortal";
            HTicketID = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.REFERENCENUMBER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            SubjectCase = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            LanguageDE = (Int16)0;
            MailTo = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailTo").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            var loopEnd4 = _.NUM(_.LEN(MailTo));
            var loopStart4 = _.NUM((Int16)1, loopEnd4);
            if (_.StrictLTE(loopStart4, loopEnd4))
            {
                for (z = loopStart4; _.StrictLTE(z, loopEnd4); z = _.ADD(z, (Int16)1))
                {
                    if (_.IF(_.EQ(_.NullableSTR(_.MID(MailTo, z, (Int16)1)), "@")))
                    {
                        CounterEmpf = _.ADD(CounterEmpf, (Int16)1);
                    }
                }
            }
            if (_.IF(_.EQ(_.ISOBJECT(_env.hlcaller), true)))
            {
                surname = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.PersonSurname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                letteraddress = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.ShortLetterAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                language = _.VAL(_.CALL(this, _env.hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.Language").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.NOTEQ(_.NullableSTR(language), "LanguageGerman")))
                {
                    LanguageDE = (Int16)(-1);
                }
                else
                {
                    LanguageDE = (Int16)1;
                }
            }
            else
            {
                surname = "Unbekannt/Unknown";
            }
            editor = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("SUINFO.EDITOR").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            //----------------------------------------------------------------------------------------------------------
            //M.Rettig, 14.05.2012 - SU-Email als HTML-Vorschau
            if (_.IF(_.EQ(_.NullableSTR(Status), "IncidentStatusClosed")))
            {
                ForReading = (Int16)1;
                ForWriting = (Int16)2;
                ForAppending = (Int16)8;

                OriginDescr = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                MailBody = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                //Deutsche Werte
                if (_.IF(_.GT(_.NullableNUM(LanguageDE), (Int16)0)))
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Herr/Frau";
                    }

                    //Konstante Werte deutsch setzen
                    TTicketID = "Ticketnummer";
                    TStatus = "Status";
                    HStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)7).Val((Int16)0).Ref(LastSUIdx, v44 => { LastSUIdx = v44; }).Val((Int16)0)));
                    TEditor = "Bearbeiter";
                    TSubject = "Betreff:";
                    if (_.IF(_.GT(_.NullableNUM(CounterEmpf), (Int16)1)))
                    {
                        Anrede = "Sehr geehrte ";
                        surname = "Damen und Herren";
                    }
                    else
                    {
                        Anrede = _.CONCAT("Sehr geehrte(r) ", _.CSTR(letteraddress));
                    }
                    TSolution = "Lösung:";
                    TBeschr = "Ticket-Beschreibung:";
                    TComplimentary = "Mit freundlichen Grüßen,";
                    TSignature = "Ihr Team IT + Prozesse";
                    TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!";
                    Creationdate = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    subject = _.CONCAT("Lösung zur IT Service Desk Anfrage ", " [#");
                    subject = _.CONCAT(subject, HTicketID, "]", " vom ", Datum);
                    TIntroduction = "Wir möchten Ihnen folgende Lösung übermitteln:";
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Mrs./Ms./Mr.";
                    }

                    //Konstante Werte englisch setzen
                    TTicketID = "Ticket number";
                    TStatus = "Status";
                    HStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)9).Val((Int16)0).Ref(LastSUIdx, v45 => { LastSUIdx = v45; }).Val((Int16)0)));
                    TEditor = "Editor";
                    TSubject = "Subject:";
                    if (_.IF(_.GT(_.NullableNUM(CounterEmpf), (Int16)1)))
                    {
                        Anrede = "Dear ";
                        surname = "Sir or Madam";
                    }
                    else
                    {
                        Anrede = _.CONCAT("Dear ", _.CSTR(letteraddress));
                    }
                    TSolution = "Solution:";
                    TBeschr = "Ticket-Description:";
                    TComplimentary = "Best regards,";
                    TSignature = "Your support team IT + Processes";
                    TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!";
                    Creationdate = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)9).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    subject = _.CONCAT("Your support request from ", Datum, " with the reference no. [#");
                    subject = _.CONCAT(subject, HTicketID, "]");
                    TIntroduction = "We deliver to you the following solution description:";
                }
                MailBody = _.REPLACE(MailBody, VBScriptConstants.vbCrLf, "<br>");
                OriginDescr = _.REPLACE(OriginDescr, VBScriptConstants.vbCrLf, "<br>");
                fso = _.OBJ(_.CREATEOBJECT("Scripting.FileSystemObject"));
                //Öffnet das File zum lesen
                f = _.OBJ(_.CALL(this, fso, "OpenTextFile", _.ARGS.Val("C:\\TRUMPF\\helpline\\Emailtemplate.html").Val(ForReading)));
                //Liest alle Daten in die Variable BodyText
                BodyText = _.VAL(_.CALL(this, f, "ReadAll"));
                BodyText = _.REPLACE(BodyText, "[$NoticeTop$]", TNoticeTop);
                BodyText = _.REPLACE(BodyText, "[$Ticket-ID_Titel$]", TTicketID);
                BodyText = _.REPLACE(BodyText, "[$TicketID$]", HTicketID);
                BodyText = _.REPLACE(BodyText, "[$Ticketstatus_Titel$]", TStatus);
                BodyText = _.REPLACE(BodyText, "[$Ticketstatus$]", HStatus);
                BodyText = _.REPLACE(BodyText, "[$Editor_Titel$]", TEditor);
                BodyText = _.REPLACE(BodyText, "[$Editor$]", editor);
                BodyText = _.REPLACE(BodyText, "[$CaseSubject_Titel$]", TSubject);
                BodyText = _.REPLACE(BodyText, "[$CaseSubject$]", SubjectCase);
                BodyText = _.REPLACE(BodyText, "[$LinktoCase_Titel$]", HLinkToCase);
                BodyText = _.REPLACE(BodyText, "[$Salutation$]", Anrede);
                BodyText = _.REPLACE(BodyText, "[$LastnameUser$]", _.CSTR(surname));
                BodyText = _.REPLACE(BodyText, "[$Introduction$]", TIntroduction);
                BodyText = _.REPLACE(BodyText, "[$CaseSolution_Titel$]", TSolution);
                BodyText = _.REPLACE(BodyText, "[$CaseSolution$]", MailBody);
                BodyText = _.REPLACE(BodyText, "[$CaseDescription_Titel$]", TBeschr);
                BodyText = _.REPLACE(BodyText, "[$CaseDescription$]", OriginDescr);
                BodyText = _.REPLACE(BodyText, "[$ComplimentaryClose$]", TComplimentary);
                BodyText = _.REPLACE(BodyText, "[$Signature$]", TSignature);
                BodyText = _.REPLACE(BodyText, "[$NoticeBottom$]", TNoticeBottom);
                //Schließt das File
                _.CALL(this, f, "Close");
                f = VBScriptConstants.Nothing;
                fso = VBScriptConstants.Nothing;
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(BodyText, v46 => { BodyText = v46; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(BodyText, v47 => { BodyText = v47; }));
            }
            else
            {
                DiagnText = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("EmailSUAttribute.EmailBody.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.EQ(_.NullableNUM(LanguageDE), (Int16)1)))
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Herr/Frau";
                    }

                    //Konstante Werte deutsch setzen
                    TTicketID = "Ticketnummer";
                    TStatus = "Status";
                    HStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)7).Val((Int16)0).Ref(LastSUIdx, v48 => { LastSUIdx = v48; }).Val((Int16)0)));
                    TEditor = "Bearbeiter";
                    TSubject = "Betreff:";
                    if (_.IF(_.GT(_.NullableNUM(CounterEmpf), (Int16)1)))
                    {
                        Anrede = "Sehr geehrte ";
                        surname = "Damen und Herren";
                    }
                    else
                    {
                        Anrede = _.CONCAT("Sehr geehrte(r) ", _.CSTR(letteraddress));
                    }
                    TDiagnosis = "Zwischenbescheid";
                    TResubTime = "Wiedervorlagedatum:";
                    TComplimentary = "Mit freundlichen Grüßen,";
                    TSignature = "Ihr Team IT + Prozesse";
                    TNoticeTop = "Bei Rückfragen antworten Sie bitte auf diese Email und verändern Sie den Betreff NICHT!";

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    Creationdate = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    ResubmissionTime = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESUBMISSIONTIME").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.NOTEQ(_.NullableSTR(ResubmissionTime), "")))
                    {
                        if (_.IF(_.GT(_.NullableNUM(_.DATEDIFF("d", _.NOW(), ResubmissionTime)), (Int16)0)))
                        {
                            //If ResubmissionTime > Now Then
                            ResubDatum = _.VAL(_.MID(ResubmissionTime, (Int16)1, (Int16)10));
                        }
                        else
                        {
                            ResubDatum = "";
                        }
                    }
                    subject = _.CONCAT("Zwischenbescheid zur IT Service Desk Anfrage ", " [#");
                    subject = _.CONCAT(subject, HTicketID, "]", " vom ", Datum);
                    TIntroduction = "Wir möchten Ihnen folgende Nachricht übermitteln:";
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Mrs./Ms./Mr.";
                    }

                    //Konstante Werte englisch setzen
                    TTicketID = "Ticket number";
                    TStatus = "Status";
                    HStatus = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.IncidentStatus").Val((Int16)9).Val((Int16)0).Ref(LastSUIdx, v49 => { LastSUIdx = v49; }).Val((Int16)0)));
                    TEditor = "Editor";
                    TSubject = "Subject:";
                    if (_.IF(_.GT(_.NullableNUM(CounterEmpf), (Int16)1)))
                    {
                        Anrede = "Dear ";
                        surname = "Sir or Madam";
                    }
                    else
                    {
                        Anrede = _.CONCAT("Dear ", _.CSTR(letteraddress));
                    }

                    TDiagnosis = "Intermediate Reply";
                    TResubTime = "Resubmissiontime:";
                    TComplimentary = "Best regards,";
                    TSignature = "Your support team IT + Processes";
                    TNoticeTop = "If you have a question or information regarding this ticket please reply to this email and do not change the subject!";

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    Creationdate = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)9).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    ResubmissionTime = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESUBMISSIONTIME").Val((Int16)9).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.NOTEQ(_.NullableSTR(ResubmissionTime), "")))
                    {
                        if (_.IF(_.GT(_.NullableNUM(_.DATEDIFF("d", _.NOW(), ResubmissionTime)), (Int16)0)))
                        {
                            //If ResubmissionTime > Now Then
                            ResubDatum = _.VAL(_.MID(ResubmissionTime, (Int16)1, (Int16)10));
                        }
                        else
                        {
                            ResubDatum = "";
                        }
                    }
                    subject = _.CONCAT("Your support request from ", Datum, " with the reference no. [#");
                    subject = _.CONCAT(subject, HTicketID, "]");
                    TIntroduction = "We deliver to you the following processing description:";
                }

                //Const ForReading = 1, ForWriting = 2, ForAppending = 8
                fso = _.OBJ(_.CREATEOBJECT("Scripting.FileSystemObject"));
                //Öffnet das File zum lesen
                f = _.OBJ(_.CALL(this, fso, "OpenTextFile", _.ARGS.Val("C:\\TRUMPF\\helpLine\\IntermediateReply.html").Val(ForReading)));
                //Liest alle Daten in die Variable BodyText
                BodyText = _.VAL(_.CALL(this, f, "ReadAll"));
                BodyText = _.REPLACE(BodyText, "[$NoticeTop$]", TNoticeTop);
                BodyText = _.REPLACE(BodyText, "[$Ticket-ID_Titel$]", TTicketID);
                BodyText = _.REPLACE(BodyText, "[$TicketID$]", HTicketID);
                BodyText = _.REPLACE(BodyText, "[$Ticketstatus_Titel$]", TStatus);
                BodyText = _.REPLACE(BodyText, "[$Ticketstatus$]", HStatus);
                BodyText = _.REPLACE(BodyText, "[$Editor_Titel$]", TEditor);
                BodyText = _.REPLACE(BodyText, "[$Editor$]", editor);
                BodyText = _.REPLACE(BodyText, "[$CaseSubject_Titel$]", TSubject);
                BodyText = _.REPLACE(BodyText, "[$CaseSubject$]", SubjectCase);
                BodyText = _.REPLACE(BodyText, "[$LinktoCase_Titel$]", HLinkToCase);
                BodyText = _.REPLACE(BodyText, "[$Salutation$]", Anrede);
                BodyText = _.REPLACE(BodyText, "[$LastnameUser$]", _.CSTR(surname));
                BodyText = _.REPLACE(BodyText, "[$Introduction$]", TIntroduction);
                BodyText = _.REPLACE(BodyText, "[$CaseInformation_Titel$]", TDiagnosis);
                BodyText = _.REPLACE(BodyText, "[$CaseInformation$]", DiagnText);
                if (_.IF(_.NOTEQ(_.NullableSTR(ResubDatum), "")))
                {
                    BodyText = _.REPLACE(BodyText, "[$ResubmissionTime_Titel$]", TResubTime);
                    BodyText = _.REPLACE(BodyText, "[$ResubmissionTime$]", ResubDatum);
                }
                else
                {
                    BodyText = _.REPLACE(BodyText, "[$ResubmissionTime_Titel$]", "");
                    BodyText = _.REPLACE(BodyText, "[$ResubmissionTime$]", "");
                }
                BodyText = _.REPLACE(BodyText, "[$ComplimentaryClose$]", TComplimentary);
                BodyText = _.REPLACE(BodyText, "[$Signature$]", TSignature);
                //Schließt das File
                _.CALL(this, f, "Close");
                f = VBScriptConstants.Nothing;
                fso = VBScriptConstants.Nothing;
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.RAWTEXT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(BodyText, v50 => { BodyText = v50; }));
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("CaseGeneral.SummaryHTML.TEXTVALUE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(BodyText, v51 => { BodyText = v51; }));
            }

        }

        public void ButtonSaveKW_Click()
        {
            object isreserved = null;
            object agent = null;
            object personid = null;
            object keywordid = null;
            object cn1 = null; /* Undeclared in source */
            object rs_person = null; /* Undeclared in source */
            object cn = null; /* Undeclared in source */
            object rs_kw = null; /* Undeclared in source */
            isreserved = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(isreserved), "")))
            {
                _.MSGBOX("Bitte zuerst das Ticket reservieren.");
            }
            else
            {

                agent = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("CASEINFO.RESERVEDBY").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));

                //Datenbankverbindung zu helpline_replication
                cn1 = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                _.SET("Provider=SQLOLEDB.1;Password=helplinereplication;Persist Security Info=True;User ID=helplinereplication;Initial Catalog=helpline_replication;Data Source=srv01itsm2", this, cn1, "ConnectionString");
                _.SET((Int16)10, this, cn1, "ConnectionTimeout");
                _.CALL(this, cn1, "Open");

                //Keyword einlesen und in Datenbank ablegen
                keywordid = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1)));
                if (_.IF(_.GT(_.NullableNUM(_.CDBL(keywordid)), (Int16)0)))
                {
                    //Personid über AgentID ermitteln
                    rs_person = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    rs_person = _.OBJ(_.CALL(this, cn1, "Execute", _.ARGS.Val(_.CONCAT("Select personid from vw_Agent_Emplkeyword where agentid = ", _.CSTR(agent)))));
                    personid = _.VAL(_.CALL(this, _.CALL(this, rs_person, "fields", _.ARGS.Val("personid")), "value"));
                    _.CALL(this, rs_person, "close");

                    //Datenbankverbindung zu helpline_data
                    cn = _.OBJ(_.CREATEOBJECT("ADODB.Connection"));
                    _.SET("Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2", this, cn, "ConnectionString");
                    _.SET((Int16)10, this, cn, "ConnectionTimeout");
                    _.CALL(this, cn, "Open");
                    //Keyword schreiben
                    rs_kw = _.OBJ(_.CREATEOBJECT("ADODB.Recordset"));
                    rs_kw = _.OBJ(_.CALL(this, cn, "Execute", _.ARGS.Val(_.CONCAT("Update dbo.emplkeywords set keyword = ", _.CDBL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("Keywords.Keyword").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1))), " where personid = ", _.CSTR(personid)))));
                    //Datenbank schließen
                    //rs_kw.close
                    _.CALL(this, cn, "close");
                    cn = VBScriptConstants.Nothing;
                }
                else
                {
                    _.MSGBOX("Please select a keyword first.");
                }

                //Datenbankverbindung zu helpline_replication schließen
                _.CALL(this, cn1, "close");
                cn1 = VBScriptConstants.Nothing;

            }

        }

        public void EditSubjectCase_ondatachange()
        {
            object Text = null;
            if (_.IF(_.INSTR((Int16)1, _.CALL(this, _env.EditSubjectCase, "Text"), "Notfalltransport_SAP", VBScriptConstants.vbTextCompare)))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }

            if (_.IF(_.INSTR((Int16)1, _.CALL(this, _env.EditSubjectCase, "Text"), "Systemänderbarkeit_SAP", VBScriptConstants.vbTextCompare)))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }

            if (_.IF(_.INSTR((Int16)1, _.CALL(this, _env.EditSubjectCase, "Text"), "#Prio 1 Incident# ", VBScriptConstants.vbTextCompare)))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }
            if (_.IF(_.INSTR((Int16)1, _.CALL(this, _env.EditSubjectCase, "Text"), "Debugg_Modus_SAP", VBScriptConstants.vbTextCompare)))
            {
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.CaseProblem, "Disabled");
                _.SET(false, this, _env.ComboBoxEmailCaller, "Disabled");
                _.SET(false, this, _env.CaseDiagnosis, "Disabled");
                _.SET(false, this, _env.KeywordTree, "Disabled");
                _.SET(false, this, _env.Attachment, "Disabled");
                _.SET(false, this, _env.CaseAttributes, "Disabled");
                _.SET(false, this, _env.ComboIncidentStatus, "Disabled");
            }

        }

        public void ButtonActionItemsAdd_Click()
        {
            object textdata = null;
            object texttemp = null;
            if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, _env.TextBoxActionItemsInput, "Text")), "")))
            {
                _.MSGBOX("Input value is missing.");
            }
            else
            {
                texttemp = _.VAL(_.CALL(this, _env.TextBoxActionItemsInput, "Text"));
                textdata = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("IncidentAttribute.ActionItems").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.NOT(_.EQ(_.NullableSTR(textdata), ""))))
                {
                    textdata = _.CONCAT(textdata, _.CHR((Int16)10), texttemp);
                }
                else
                {
                    textdata = _.VAL(texttemp);
                }
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ActionItems").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(textdata, v52 => { textdata = v52; }));
            }

        }

        public void ButtonActionItemsDel_Click()
        {
            object delete = null;
            delete = _.VAL(_.CALL(this, _, "MSGBOX", _.ARGS.Val("Delete all action items permanently?").Val((Int16)4).Val("Delete Action Items")));
            if (_.IF(_.EQ(_.NullableNUM(delete), (Int16)6)))
            {
                _.CALL(this, _env.hlobj, "SetValue", _.ARGS.Val("IncidentAttribute.ActionItems").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val(""));
            }

        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object Asset { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object Attachment { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object AttachmentControlATTACHMENT { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object AttachmentControlATTACHMENT1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object AttachmentControlATTACHMENT2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object b_template_change { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object b_template_delete { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object b_template_load { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object b_template_save { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonActionItemsAdd { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonActionItemsDel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonCC { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonDiscovery { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonEmailPreview { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonResetTo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSaveKW { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSCCMRemote { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSearchMail { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSetAgent { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSetKW { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonShowOverView { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonTo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object CaseAttributes { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object CaseDiagnosis { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object CaseProblem { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object cb_template_load { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object CheckBoxPUBLISHED { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object CheckBoxStandby1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboBoxEmailCaller { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboBoxEmailSearchResult { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboFunctionalRange { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboImpact { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboIncidentStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboLevel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboPriority { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboProductionalRelevanz { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboRequestType { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComboVIPStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComplexText1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComplexTextEmailBody { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ComplexTextSummaryHTML { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object DateTimeControlFailureend1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object DateTimeControlFailurestart1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditAssetModel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditCinumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditDiagnosis { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditEmailAddress { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditGivenName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditHostname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditIncStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditOrganisation { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditPhoneNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditProblem { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditResubmissionTime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditSubjectCase { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object EditSurname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object Formular1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object gb_Templater { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object gb_vorgaenge_background { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object gb_vorgaenge_background1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBox2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxChanges1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxChangesTree { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxChangesTree1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxEmail { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxNotifications { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlcaller { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlobj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlSession { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object InfoArea { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object InfoReferenceNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object InfoRegistrationtime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object KeywordTree { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object l_template { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object l_templateID { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelActionItems { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelAffectedArea { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelAssetModel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelCinumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEDITOR { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailAddress { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailBody { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailCaller1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailCC { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailFrom { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailSearchName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailSearchResult { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailSubject { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelEmailTo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelFailureend1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelFailurestart1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelFunctionalRange { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelGivenName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelHostname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelImpact { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelInfoRegistrationtime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelLevel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelPhoneNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelPriority { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelProductionalRelevanz { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelREFERENCENUMBER { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelRequestType { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelResubmissionTime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelRoomNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelSubjectCase { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelSurname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object LabelVIPStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object Person { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object SearchAsset { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object SearchCaller { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object SUNavigator1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object SUNavigator2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabCombinedCases { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabCustomer { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableAssetCases { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableCaseFolder1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableControl1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableControl2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableDelegatedCases1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableMainUser { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableParentCase1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableProduct { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TableRequesterCases { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabOverview { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageChanges { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageEmail { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageSolution { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabRequesterCases { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object Tabs { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxActionItemsData { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxActionItemsInput { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxAffectedArea { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxCREATIONTIME { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEDITOR { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEmailCC { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEmailFrom { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEmailSearchName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEmailSubject { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxEmailTo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSolutionText { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TreeKeyword { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TreeKeywordOrga { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TreeSelControlKeywordNotification { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TreeSelControlKeywordtask1 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}