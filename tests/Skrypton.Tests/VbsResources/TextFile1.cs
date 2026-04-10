using System;
using System.Collections;
using Skrypton.RuntimeSupport;
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
            //---------------------------------------------------------------

            //---------------------------------------------------------------
            //SACM
            //----------------------------------------------------------------------------------------------------------
            //Globale Konstanten für freie Assoziationsdefinitionen

            //----------------------------------------------------------------------------------------------------------

            //---------------------------------------------------------------------
            // Check whether the agent (contact) is allowed to make
            // changes/modifications/create new entities of any objectdefinition based
            // on the InternalMIGPartnerID of the contact and the MIGPartnerID of the object

            //---------------------------------------------------------------------
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
            HLASC_SoftwareLicenseFolderView = "LicenseFolderView";
            HLASC_SoftwareLicenseGroupView = "LicenseGroupView";
            HLASC_Software2Computer = "Software2Computer";
        }
        internal object HLASC_SoftwareLicenseFolderView { get; set; }
        internal object HLASC_SoftwareLicenseGroupView { get; set; }
        internal object HLASC_Software2Computer { get; set; }
        //---------------------------------------------------------------
        //Diese Funktion ermittelt den Standard-Eintrag zum angegebenen Attribut aus
        //dem Dictionary.
        //Wenn der Parameter "GetAll" auf False steht wird als Rückgabewert für die Funktion
        //ebenfalls "False" ausgegben, wenn mehr als ein Standardeintrag gefunden wird.
        //Wenn für den Parameter "True" angeben wird, prüft die Funktion ob es tatsächlich
        //nur einen Standard-Eintrag gibt, sonst "False".
        public object GetCommunicationDefault(ref object hlContext, ref object hlObject, ref object dict, ref object GetAll)
        {
            object GetCommunicationDefault_retVal = null;
            object ItemCount = null;
            object strValue = null;
            object ItemIDs = null;
            object Item = null;
            object defItem = null;
            GetCommunicationDefault_retVal = false;
            ItemCount = (Int16)0;
            strValue = "";

            ItemIDs = "";
            object dict_vref = dict;
            try
            {
                ItemIDs = _.VAL(_.CALLm1argp(this, _.NnO(hlObject, "hlObject"), "GetContentIDs", _.ARGS.RefIfArray(dict_vref, _.ARGS.Val("Compound")).Val((Int16)0)));
            }
            finally { dict = dict_vref; }

            Item = (Int16)0;
            var enumerationContent = _.ENUMERABLE(ItemIDs).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                Item = enumerationContent.Current;
                defItem = false;
                object hlContext_vref = hlContext, hlObject_vref = hlObject, dict_vref2 = dict;
                try
                {
                    defItem = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(hlContext_vref, v => { hlContext_vref = v; }).Ref(hlObject_vref, v2 => { hlObject_vref = v2; }).RefIfArray(dict_vref2, _.ARGS.Val("Default")).Ref(Item, v3 => { Item = v3; }).Val((Int16)0)));
                }
                finally { hlContext = hlContext_vref; hlObject = hlObject_vref; dict = dict_vref2; }
                if (_.IF(_.EQ(_.CBOOL(defItem), true)))
                {
                    ItemCount = _.ADD(ItemCount, (Int16)1);
                    object dict_vref3 = dict;
                    try
                    {
                        strValue = _.VAL(_.CALLm1argp(this, _.NnO(hlObject, "hlObject"), "GetValue", _.ARGS.RefIfArray(dict_vref3, _.ARGS.Val("Value")).Val((Int16)0).Ref(Item, v4 => { Item = v4; }).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { dict = dict_vref3; }
                    if (_.IF(_.EQ(_.CBOOL(GetAll), false)))
                    {
                        break;
                    }
                }
            }
            if (_.IF(_.GT(_.NullableNUM(ItemCount), (Int16)1)))
            {
                GetCommunicationDefault_retVal = false;
                return GetCommunicationDefault_retVal;
            }
            else
            {
                GetCommunicationDefault_retVal = true;
                _.SETm0a1(this, _.NnO(dict, "dict"), "DefValue", _.VAL(strValue));
            }
            return GetCommunicationDefault_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
        public void Trace(ref object hlContext, ref object text)
        {
            object text_vref = text;
            try
            {
                _.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "trace", _.ARGS.Val((Int16)1).Ref(text_vref, v5 => { text_vref = v5; }));
            }
            finally { text = text_vref; }
        }
        //---------------------------------------------------------------
        //Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
        public void SetPersonInformation(ref object hlContext, ref object hlObject, ref object dict)
        {
            object AttrDef = null;
            object strAttrValue = null;
            //Aus dem Dictionary wird das Attribut und der dazugehörige Wert ermittelt.
            AttrDef = "";
            AttrDef = _.CONCAT("PersonInformation.", _.CALLm0argp(this, _.NnO(dict, "dict"), _.ARGS.Val("PersInfoAttr")));

            strAttrValue = "";
            strAttrValue = _.VAL(_.CALLm0argp(this, _.NnO(dict, "dict"), _.ARGS.Val("DefValue")));

            if (_.IF(_.EQ(_.NullableSTR(strAttrValue), "")))
            {
                strAttrValue = "-";
            }
            _.CALLm1argp(this, _.NnO(hlObject, "hlObject"), "SetValue", _.ARGS.Ref(AttrDef, v6 => { AttrDef = v6; }).Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strAttrValue, v7 => { strAttrValue = v7; }));
        }
        //---------------------------------------------------------------
        public object IsHLObject(ref object hlContext, ref object hlObject)
        {
            object IsHLObject_retVal = null;
            //	Trace hlContext, "IsObject " & IsObject(hlObject)
            //	Trace hlContext, "IsNull " & IsNull(hlObject)
            //	Trace hlContext, "IsEmpty " & IsEmpty(hlObject)
            //	Trace hlContext, "Leerstring "
            //	Trace hlContext, "Leerstring " & hlObject = ""
            object hlContext_vref2 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(hlContext_vref2, v8 => { hlContext_vref2 = v8; }).Val(_.CONCAT("Type ", _.VARTYPE(hlObject))));
            }
            finally { hlContext = hlContext_vref2; }
            IsHLObject_retVal = _.VAL(_.ANDe2(_.CBOOL(_.EQ(_.ISOBJECT(hlObject), true)) && _.CBOOL(_.EQ(_.IS(hlObject, VBScriptConstants.Nothing), false))));
            return IsHLObject_retVal;
        }
        //-------------------------------------------------------------------
        public object GetBaseType(ref object hlContext, ref object hlObject)
        {
            return _.VAL(_.CALLm1v5(this, _.NnO(hlObject, "hlObject"), "GetValue", "HLOBJECTINFO.BASETYPE", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
        }
        //---------------------------------------------------------------
        //Dies ist eine rekursive Function zum ermitteln der Organisationshierarchie,
        //ausgehend vom der ersten OU überhalb einer Person.
        //Die Variable "strOrgUnits" ist der Out-Parameter der Function.
        public object GetPersonOrganisation(ref object hlContext, ref object hlOrgUnit, ref object strOrgUnits)
        {
            object GetPersonOrganisation_retVal = null;
            object retval = null;
            object NextOrgUnit = null;
            object orgaType = null;
            GetPersonOrganisation_retVal = (Int16)0;
            retval = (Int16)0;

            //Wenn noch keine OU ermittelt wurde, wird der Name der ersten OU eingetragen.
            //Andernfalls, wird jede weitere OU einfach angehangen.
            if (_.IF(_.EQ(_.NullableSTR(strOrgUnits), "")))
            {
                strOrgUnits = _.VAL(_.CALLm1v5(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            }
            else
            {
                strOrgUnits = _.CONCAT(strOrgUnits, ", ", _.CALLm1v5(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            }

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            orgaType = "";
            orgaType = _.VAL(_.CALLm1v0(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Division")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetItems", 65536, (Int16)0, (Int16)0, "CompanyView"));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Site")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetItems", 65536, (Int16)0, (Int16)0, "Site2Company"));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Company")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, _.NnO(hlOrgUnit, "hlOrgUnit"), "GetItems", 65536, (Int16)0, (Int16)0, "Company2Company"));
            }

            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.ISARRAY(NextOrgUnit)))
            {
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextOrgUnit)), (Int16)0)))
                {
                    object hlContext_vref3 = hlContext, strOrgUnits_vref = strOrgUnits;
                    try
                    {
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(hlContext_vref3, v9 => { hlContext_vref3 = v9; }).RefIfArray(NextOrgUnit, _.ARGS.Val((Int16)0)).Ref(strOrgUnits_vref, v10 => { strOrgUnits_vref = v10; })));
                    }
                    finally { hlContext = hlContext_vref3; strOrgUnits = strOrgUnits_vref; }
                }
                else
                {
                    return GetPersonOrganisation_retVal;
                }
            }
            return GetPersonOrganisation_retVal;
        }
        //---------------------------------------------------------------
        //Über diese Function wird für ein Flag Attribut immer der Wert
        //True oder False ausgegeben.
        public object GetFlagValue(ref object hlContext, ref object hlObject, ref object hlattribute, ref object hlcontentid, ref object hlsuid)
        {
            object GetFlagValue_retVal = null;
            object hlattribute_vref = hlattribute, hlcontentid_vref = hlcontentid, hlsuid_vref = hlsuid;
            try
            {
                GetFlagValue_retVal = _.VAL(_.CALLm1argp(this, _.NnO(hlObject, "hlObject"), "GetValue", _.ARGS.Ref(hlattribute_vref, v11 => { hlattribute_vref = v11; }).Val((Int16)0).Ref(hlcontentid_vref, v12 => { hlcontentid_vref = v12; }).Ref(hlsuid_vref, v13 => { hlsuid_vref = v13; }).Val((Int16)0)));
            }
            finally { hlattribute = hlattribute_vref; hlcontentid = hlcontentid_vref; hlsuid = hlsuid_vref; }
            if (_.IF(_.EQ(_.NullableSTR(GetFlagValue_retVal), "")))
            {
                GetFlagValue_retVal = false;
            }
            return GetFlagValue_retVal;
        }
        //-------------------------------------------------------------------
        //Diese Function ermitellt eine Fehlermeldung aus dem helpLine
        //Wörterbuch ohne Parameter.
        public object GetErrMsg0(ref object hlContext, ref object LocaleID, ref object ErrCode)
        {
            object GetErrMsg0_retVal = null;
            object strErrMsg = null;
            GetErrMsg0_retVal = "";

            strErrMsg = "";
            object ErrCode_vref = ErrCode, LocaleID_vref = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetTranslation", _.ARGS.Ref(ErrCode_vref, v14 => { ErrCode_vref = v14; }).Ref(LocaleID_vref, v15 => { LocaleID_vref = v15; })));
            }
            finally { ErrCode = ErrCode_vref; LocaleID = LocaleID_vref; }
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbNewLine, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg0_retVal = _.REPLACE(strErrMsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg0_retVal;
        }
        //Das Script ermittelt auf Basis der ersten übergeordneten OU den gesamten Pfad bis zur Firma oder Konzern
        //und speichert diesen in das Hilfsattribut PersonInformation.PersonOrganisation.
        //This script detects the entire path based on the first parent OU up to the company or holding
        //and saves them into the attribute PersonInformation.PersonOrganisation.
        public void SetPersonOrganization(ref object hlContext, ref object hlPerson, ref object dict)
        {
            object FirstOrgUnit = null;
            object rsltOrgUnit = null;
            object retval = null;
            object strOrgUnits = null;
            FirstOrgUnit = VBScriptConstants.Nothing;
            FirstOrgUnit = _.OBJ(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "GetRelatedObject"));

            bool ifResult;
            object hlContext_vref4 = hlContext;
            try
            {
                ifResult = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(hlContext_vref4, v18 => { hlContext_vref4 = v18; }).Ref(FirstOrgUnit, v19 => { FirstOrgUnit = v19; })), true));
            }
            finally { hlContext = hlContext_vref4; }
            if (ifResult)
            {
                if (_.IF(_.ANDe2(_.CBOOL(_.NOTEQ(_.NullableSTR(_.CALLm1v0(this, _.NnO(FirstOrgUnit, "FirstOrgUnit"), "GetType")), "Company")) && _.CBOOL(_.NOTEQ(_.NullableSTR(_.CALLm1v0(this, _.NnO(FirstOrgUnit, "FirstOrgUnit"), "GetType")), "Division")))))
                {
                    FirstOrgUnit = VBScriptConstants.Nothing;
                }
            }

            bool ifResult2;
            object hlContext_vref5 = hlContext;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(hlContext_vref5, v22 => { hlContext_vref5 = v22; }).Ref(FirstOrgUnit, v23 => { FirstOrgUnit = v23; })), false));
            }
            finally { hlContext = hlContext_vref5; }
            if (ifResult2)
            {
                rsltOrgUnit = "";
                rsltOrgUnit = _.VAL(_.CALLm1v4(this, _.NnO(hlPerson, "hlPerson"), "GetItems", 65536, (Int16)0, (Int16)0, "Person2Organization"));
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltOrgUnit)), (Int16)0)))
                {
                    FirstOrgUnit = _.OBJ(_.CALLm0argp(this, _.NnO(rsltOrgUnit, "rsltOrgUnit"), _.ARGS.Val((Int16)0)));
                }
            }

            bool ifResult3;
            object hlContext_vref6 = hlContext;
            try
            {
                ifResult3 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(hlContext_vref6, v26 => { hlContext_vref6 = v26; }).Ref(FirstOrgUnit, v27 => { FirstOrgUnit = v27; })), true));
            }
            finally { hlContext = hlContext_vref6; }
            if (ifResult3)
            {
                bool ifResult4;
                object hlContext_vref7 = hlContext;
                try
                {
                    ifResult4 = _.IF(_.EQ(_.NullableSTR(_.CALLm1argp(this, _outer, "GetBaseType", _.ARGS.Ref(hlContext_vref7, v30 => { hlContext_vref7 = v30; }).Ref(FirstOrgUnit, v31 => { FirstOrgUnit = v31; }))), "ORGANISATION"));
                }
                finally { hlContext = hlContext_vref7; }
                if (ifResult4)
                {
                    retval = "";
                    strOrgUnits = "";
                    object hlContext_vref8 = hlContext;
                    try
                    {
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(hlContext_vref8, v32 => { hlContext_vref8 = v32; }).Ref(FirstOrgUnit, v33 => { FirstOrgUnit = v33; }).Ref(strOrgUnits, v34 => { strOrgUnits = v34; })));
                    }
                    finally { hlContext = hlContext_vref8; }

                    _.SETm0a1(this, _.NnO(dict, "dict"), "DefValue", _.VAL(strOrgUnits));
                    _.SETm0a1(this, _.NnO(dict, "dict"), "PersInfoAttr", "PersonOrganization");
                    object hlContext_vref9 = hlContext, hlPerson_vref = hlPerson, dict_vref4 = dict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "SetPersonInformation", _.ARGS.Ref(hlContext_vref9, v35 => { hlContext_vref9 = v35; }).Ref(hlPerson_vref, v36 => { hlPerson_vref = v36; }).Ref(dict_vref4, v37 => { dict_vref4 = v37; }));
                    }
                    finally { hlContext = hlContext_vref9; hlPerson = hlPerson_vref; dict = dict_vref4; }
                }
            }
        }
        //----------------------------------------------------------------------------------------------------------
        //Prozedur füllt die Umzugshistorie für das entsprechende Objekt
        public void SetAssetHistory(ref object hlContext, ref object hlObjectA, ref object hlObjectB, ref object created)
        {
            object productDefName = null;
            object agentID = null;
            object contentID = null;
            object personOfAgent = null;
            object personName = null;
            object orgUnitName = null;
            object strErrMsg = null;

            productDefName = _.VAL(_.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "GetType", _.ARGS.ForceBrackets()));

            if (_.IF(_.ANDe2(_.CBOOL(_.NOTEQ(_.NullableSTR(productDefName), "Software")) && _.CBOOL(_.NOTEQ(_.NullableSTR(productDefName), "SoftwareLicence")))))
            {
                contentID = _.VAL(_.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "GenerateContentID", _.ARGS.ForceBrackets()));
                agentID = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetAgentID", _.ARGS.ForceBrackets()));
                orgUnitName = _.VAL(_.CALLm1v5(this, _.NnO(hlObjectA, "hlObjectA"), "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                personOfAgent = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetPersonOfAgent", _.ARGS.Ref(agentID, v38 => { agentID = v38; })));
                if (_.IF(_.IS(personOfAgent, VBScriptConstants.Nothing)))
                {
                    object hlContext_vref10 = hlContext;
                    try
                    {
                        strErrMsg = _.VAL(_.CALLm1argp(this, _outer, "GetErrMsg0", _.ARGS.Ref(hlContext_vref10, v39 => { hlContext_vref10 = v39; }).Val(_.CALLm1v0(this, _.NnO(hlContext_vref10, "hlContext_vref10"), "GetLocaleID")).Val("#ERR_SETASSETHISTORY")));
                    }
                    finally { hlContext = hlContext_vref10; }
                    object hlContext_vref11 = hlContext;
                    try
                    {
                        _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(hlContext_vref11, v40 => { hlContext_vref11 = v40; }).Ref(strErrMsg, v41 => { strErrMsg = v41; }));
                    }
                    finally { hlContext = hlContext_vref11; }
                    //hlContext.abortcommand strErrMsg
                }
                else
                {
                    personName = _.VAL(_.CALLm1v5(this, _.NnO(personOfAgent, "personOfAgent"), "GetValue", "PersonGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    personName = _.CONCAT(personName, ", ");
                    personName = _.CONCAT(personName, _.CALLm1v5(this, _.NnO(personOfAgent, "personOfAgent"), "GetValue", "PersonGeneral.GivenName", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                }
                _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedBy").Val((Int16)0).Ref(contentID, v42 => { contentID = v42; }).Val((Int16)0).Ref(personName, v43 => { personName = v43; }));
                _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID").Val((Int16)0).Ref(contentID, v44 => { contentID = v44; }).Val((Int16)0).Ref(agentID, v45 => { agentID = v45; }));
                _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangeDate").Val((Int16)0).Ref(contentID, v46 => { contentID = v46; }).Val((Int16)0).Val(_.NOW()));
                _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnit").Val((Int16)0).Ref(contentID, v47 => { contentID = v47; }).Val((Int16)0).Ref(orgUnitName, v48 => { orgUnitName = v48; }));
                _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnitID").Val((Int16)0).Ref(contentID, v49 => { contentID = v49; }).Val((Int16)0).Val(_.CALLm1argp(this, _.NnO(hlObjectA, "hlObjectA"), "GetID", _.ARGS.ForceBrackets())));

                if (_.IF(_.EQ(created, true)))
                {
                    _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v50 => { contentID = v50; }).Val((Int16)0).Val("HistoryActionCreated"));
                }
                else
                {
                    _.CALLm1argp(this, _.NnO(hlObjectB, "hlObjectB"), "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v51 => { contentID = v51; }).Val((Int16)0).Val("HistoryActionDeleted"));
                }
            }
        }
        //---------------------------------------------------------------
        //Diese Function ermitellt eine Fehlermeldung aus dem helpLine
        //Wörterbuch mit einem Parameter.
        public object GetErrMsg1(ref object hlContext, ref object LocaleID, ref object ErrCode, ref object Arg1)
        {
            object GetErrMsg1_retVal = null;
            object strErrMsg = null;
            GetErrMsg1_retVal = "";

            strErrMsg = "";
            object ErrCode_vref2 = ErrCode, LocaleID_vref2 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetTranslation", _.ARGS.Ref(ErrCode_vref2, v52 => { ErrCode_vref2 = v52; }).Ref(LocaleID_vref2, v53 => { LocaleID_vref2 = v53; })));
            }
            finally { ErrCode = ErrCode_vref2; LocaleID = LocaleID_vref2; }
            strErrMsg = _.REPLACE(strErrMsg, "%1", Arg1);
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbLf, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg1_retVal = _.REPLACE(strErrMsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg1_retVal;
        }
        public object GetErrMsg2(ref object hlContext, ref object LocaleID, ref object ErrCode, ref object Arg1, ref object Arg2)
        {
            object GetErrMsg2_retVal = null;
            object strErrMsg = null;
            GetErrMsg2_retVal = "";

            strErrMsg = "";
            object ErrCode_vref3 = ErrCode, LocaleID_vref3 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetTranslation", _.ARGS.Ref(ErrCode_vref3, v54 => { ErrCode_vref3 = v54; }).Ref(LocaleID_vref3, v55 => { LocaleID_vref3 = v55; })));
            }
            finally { ErrCode = ErrCode_vref3; LocaleID = LocaleID_vref3; }
            strErrMsg = _.REPLACE(strErrMsg, "%1", Arg1);
            strErrMsg = _.REPLACE(strErrMsg, "%2", Arg2);
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbLf, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg2_retVal = _.REPLACE(strErrMsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg2_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //In dieser Funktion wird geprüft, ob es unterhalb einer Software Suite
        //bereits Lizenzumschläge mit Lizenzen gibt.
        public object GetReferenceLicenseCount(ref object hlContext, ref object hlSWFolder, ref object chkFolderOnly, ref object HLASC_SoftwareLicenseFolderView)
        {
            object GetReferenceLicenseCount_retVal = null;
            object rsltSWFolders = null;
            object SoftwareLicense = null;
            object objType = null;
            GetReferenceLicenseCount_retVal = (Int16)0;

            rsltSWFolders = "";
            SoftwareLicense = VBScriptConstants.Nothing;
            objType = "";

            //Prüfen ob es Software Lizenzobjekte/Lizenzumschläge unterhalb des Folders gibt.
            object HLASC_SoftwareLicenseFolderView_vref = HLASC_SoftwareLicenseFolderView;
            try
            {
                rsltSWFolders = _.VAL(_.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(HLASC_SoftwareLicenseFolderView_vref, v56 => { HLASC_SoftwareLicenseFolderView_vref = v56; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = HLASC_SoftwareLicenseFolderView_vref; }

            var enumerationContent2 = _.ENUMERABLE(rsltSWFolders).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                SoftwareLicense = enumerationContent2.Current;
                objType = _.VAL(_.CALLm1argp(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objType), "LicenseFolder")))
                {
                    object hlContext_vref12 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref12, v57 => { hlContext_vref12 = v57; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref12; }
                    if (_.IF(_.GT(_.NullableNUM(GetReferenceLicenseCount_retVal), (Int16)0)))
                    {
                        return GetReferenceLicenseCount_retVal;
                    }
                }
                if (_.IF(_.ANDe2(_.CBOOL(_.EQ(_.NullableSTR(objType), "SoftwareLicense")) && _.CBOOL(_.EQ(_.CBOOL(chkFolderOnly), false)))))
                {
                    object hlContext_vref13 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref13, v58 => { hlContext_vref13 = v58; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref13; }
                    if (_.IF(_.GT(_.NullableNUM(GetReferenceLicenseCount_retVal), (Int16)0)))
                    {
                        return GetReferenceLicenseCount_retVal;
                    }
                }
            }
            return GetReferenceLicenseCount_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
        //den obersten Lizenz Umschlag ermittelt. Auf dem Weg dort hin wird geprüft ob einer
        //der Lizenzumschläge eine Software Suite ist.
        public object CheckForSoftwareSuiteFolder(ref object hlContext, ref object hlParentSWFolder, ref object pDict, ref object HLASC_SoftwareLicenseFolderView)
        {
            object CheckForSoftwareSuiteFolder_retVal = null;
            object retval = null;
            object NextSWFolder = null;
            object CheckSoftwareSuite = null;
            CheckForSoftwareSuiteFolder_retVal = "";
            retval = (Int16)0;
            NextSWFolder = "";

            //Festhalten auf welcher Ebene ggf. eine Software Suite oberhalb des
            //Start Folders existiert. Die Variable muss von außen mit einem Startwert
            //initialisiert werden.
            if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SoftwareSuiteFolderLevel"))), (Int16)0), _.EQ(_.NullableSTR(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SoftwareSuiteFolderLevel"))), ""))))
            {
                _.SETm0a1(this, _.NnO(pDict, "pDict"), "SoftwareSuiteFolderLevel", (Int16)1);
            }
            else
            {
                _.SETm0a1(this, _.NnO(pDict, "pDict"), "SoftwareSuiteFolderLevel", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SoftwareSuiteFolderLevel")), (Int16)1));
            }

            //Amhand des Flags "Software Suite" festellen ob ein Lizenzumschlag als Software Suite
            //gekennzeichnet ist. Falls Ja, Name des Umschlags auslesen und Funktion abbrechen.
            CheckSoftwareSuite = false;
            object hlContext_vref14 = hlContext, hlParentSWFolder_vref = hlParentSWFolder;
            try
            {
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(hlContext_vref14, v59 => { hlContext_vref14 = v59; }).Ref(hlParentSWFolder_vref, v60 => { hlParentSWFolder_vref = v60; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = hlContext_vref14; hlParentSWFolder = hlParentSWFolder_vref; }
            if (_.IF(_.EQ(_.CBOOL(CheckSoftwareSuite), true)))
            {
                _.SETm0a1(this, _.NnO(pDict, "pDict"), "SoftwareSuiteFolder", _.VAL(_.CALLm1v5(this, _.NnO(hlParentSWFolder, "hlParentSWFolder"), "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
                return CheckForSoftwareSuiteFolder_retVal;
            }

            //Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            object HLASC_SoftwareLicenseFolderView_vref2 = HLASC_SoftwareLicenseFolderView;
            try
            {
                NextSWFolder = _.VAL(_.CALLm1argp(this, _.NnO(hlParentSWFolder, "hlParentSWFolder"), "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(HLASC_SoftwareLicenseFolderView_vref2, v61 => { HLASC_SoftwareLicenseFolderView_vref2 = v61; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = HLASC_SoftwareLicenseFolderView_vref2; }
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextSWFolder)), (Int16)0)))
            {
                object hlContext_vref15 = hlContext, pDict_vref = pDict, HLASC_SoftwareLicenseFolderView_vref3 = HLASC_SoftwareLicenseFolderView;
                try
                {
                    retval = _.VAL(_.CALLm1argp(this, _outer, "CheckForSoftwareSuiteFolder", _.ARGS.Ref(hlContext_vref15, v62 => { hlContext_vref15 = v62; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(pDict_vref, v63 => { pDict_vref = v63; }).Ref(HLASC_SoftwareLicenseFolderView_vref3, v64 => { HLASC_SoftwareLicenseFolderView_vref3 = v64; })));
                }
                finally { hlContext = hlContext_vref15; pDict = pDict_vref; HLASC_SoftwareLicenseFolderView = HLASC_SoftwareLicenseFolderView_vref3; }
            }
            else
            {
                return CheckForSoftwareSuiteFolder_retVal;
            }
            return CheckForSoftwareSuiteFolder_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //In dieser Rekursiven Funktion wird solange nach oben gegangen, bis man
        //den obersten Lizenz Umschlag ermittelt und neu berechnet hat.
        public object SetLicenseCounter(ref object hlContext, ref object hlSWFolder, ref object pDict, ref object assocName)
        {
            object SetLicenseCounter_retVal = null;
            object retval = null;
            object CheckSoftwareSuite = null;
            object CheckLicContrByServer = null;
            object NextSWFolder = null;
            object a = null;
            SetLicenseCounter_retVal = (Int16)0;
            retval = (Int16)0;

            //Dictionary Einträge initalisieren
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SoftwareLicenses", "");
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumRefLicCounter", (Int16)0);
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumInstLicCounter", (Int16)0);
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumFreeLicCounter", (Int16)0);

            //Prüfen ob es Software Lizenzobjekte unterhalb des Folders gibt.
            object assocName_vref = assocName;
            try
            {
                _.SETm0a1(this, _.NnO(pDict, "pDict"), "SoftwareLicenses", _.VAL(_.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(assocName_vref, v65 => { assocName_vref = v65; }))));
            }
            finally { assocName = assocName_vref; }

            //Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
            //Objekte gezählt werden müssen
            CheckSoftwareSuite = false;
            object hlContext_vref16 = hlContext, hlSWFolder_vref = hlSWFolder;
            try
            {
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(hlContext_vref16, v66 => { hlContext_vref16 = v66; }).Ref(hlSWFolder_vref, v67 => { hlSWFolder_vref = v67; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = hlContext_vref16; hlSWFolder = hlSWFolder_vref; }

            bool ifResult5;
            object pDict_vref2 = pDict;
            try
            {
                ifResult5 = _.IF(_.GTE(_.NullableNUM(_.UBOUND(_.CALLm0argp(this, _.NnO(pDict_vref2, "pDict_vref2"), _.ARGS.Val("SoftwareLicenses")))), (Int16)0));
            }
            finally { pDict = pDict_vref2; }
            if (ifResult5)
            {
                if (_.IF(_.EQ(_.CBOOL(CheckSoftwareSuite), false)))
                {
                    object hlContext_vref17 = hlContext, pDict_vref3 = pDict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "CalcAllLicCounter", _.ARGS.Ref(hlContext_vref17, v68 => { hlContext_vref17 = v68; }).Ref(pDict_vref3, v69 => { pDict_vref3 = v69; }));
                    }
                    finally { hlContext = hlContext_vref17; pDict = pDict_vref3; }
                }
                else
                {
                    object hlContext_vref18 = hlContext, pDict_vref4 = pDict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "CalcFolderLicCounter", _.ARGS.Ref(hlContext_vref18, v70 => { hlContext_vref18 = v70; }).Ref(pDict_vref4, v71 => { pDict_vref4 = v71; }));
                    }
                    finally { hlContext = hlContext_vref18; pDict = pDict_vref4; }
                }
            }
            //Gesatmzahl der Lizenzen in den Lizenzumschlag zurückschreiben
            object pDict_vref5 = pDict;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref5, _.ARGS.Val("SumRefLicCounter")));
            }
            finally { pDict = pDict_vref5; }
            object pDict_vref6 = pDict;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "SetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref6, _.ARGS.Val("SumInstLicCounter")));
            }
            finally { pDict = pDict_vref6; }

            //Wenn die Lizenzkontrolle durch den Applikations Server erfolgt ("Lizenzkontrolle durch Server")
            //dann die Anzahl freier Lizenzen immer auf den Wert "0" setzen.
            CheckLicContrByServer = false;
            object hlContext_vref19 = hlContext, hlSWFolder_vref2 = hlSWFolder;
            try
            {
                CheckLicContrByServer = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(hlContext_vref19, v72 => { hlContext_vref19 = v72; }).Ref(hlSWFolder_vref2, v73 => { hlSWFolder_vref2 = v73; }).Val("SoftwareLicenseFolderDetail.FlagLicenseControlledByServer").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = hlContext_vref19; hlSWFolder = hlSWFolder_vref2; }
            if (_.IF(_.EQ(_.CBOOL(CheckLicContrByServer), true)))
            {
                _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumFreeLicCounter", (Int16)0);
            }
            object pDict_vref7 = pDict;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "SetValue", _.ARGS.Val("SoftwareLicenseCounter.FreeLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref7, _.ARGS.Val("SumFreeLicCounter")));
            }
            finally { pDict = pDict_vref7; }

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            NextSWFolder = "";
            a = "";
            a = _.VAL(_.CALLm1v0(this, _.NnO(hlSWFolder, "hlSWFolder"), "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(a), "LicenseFolder")))
            {
                object assocName_vref2 = assocName;
                try
                {
                    NextSWFolder = _.VAL(_.CALLm1argp(this, _.NnO(hlSWFolder, "hlSWFolder"), "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Ref(assocName_vref2, v74 => { assocName_vref2 = v74; })));
                }
                finally { assocName = assocName_vref2; }
            }
            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextSWFolder)), (Int16)0)))
            {
                object hlContext_vref20 = hlContext, pDict_vref8 = pDict, assocName_vref3 = assocName;
                try
                {
                    retval = _.VAL(_.CALLm1argp(this, _outer, "SetLicenseCounter", _.ARGS.Ref(hlContext_vref20, v75 => { hlContext_vref20 = v75; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(pDict_vref8, v76 => { pDict_vref8 = v76; }).Ref(assocName_vref3, v77 => { assocName_vref3 = v77; })));
                }
                finally { hlContext = hlContext_vref20; pDict = pDict_vref8; assocName = assocName_vref3; }
            }
            else
            {
                return SetLicenseCounter_retVal;
            }
            return SetLicenseCounter_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        public object IsValidObject(ref object obj)
        {
            return _.VAL(_.ANDe2(_.CBOOL(_.ISOBJECT(obj)) && _.CBOOL(_.NOT(_.IS(obj, VBScriptConstants.Nothing)))));
        }
        //----------------------------------------------------------------------------------------------------------
        public void CalcAllLicCounter(ref object hlContext, ref object pDict)
        {
            object SWRefLicCounter = null;
            object SWInstCounter = null;
            object SoftwareLicense = null;
            object objType = null;
            object lstLicStatus = null;
            SWRefLicCounter = (Int16)0;
            SWInstCounter = (Int16)0;
            SoftwareLicense = VBScriptConstants.Nothing;
            objType = "";
            lstLicStatus = "";

            var enumerationContent3 = _.ENUMERABLE(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                SoftwareLicense = enumerationContent3.Current;
                objType = _.VAL(_.CALLm1argp(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objType), "SoftwareLicense")))
                {
                    lstLicStatus = _.VAL(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseDetail.LicenseStatus", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object hlContext_vref21 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref21, v78 => { hlContext_vref21 = v78; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref21; }
                        _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumRefLicCounter", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                    }
                }
                else
                {
                    if (_.IF(_.OR(_.EQ(_.NullableSTR(objType), "LicenseFolder"), _.EQ(_.NullableSTR(objType), "Software"))))
                    {
                        object hlContext_vref22 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref22, v79 => { hlContext_vref22 = v79; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref22; }
                        _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumRefLicCounter", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                        object hlContext_vref23 = hlContext;
                        try
                        {
                            SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref23, v80 => { hlContext_vref23 = v80; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.InstalledLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref23; }
                        _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumInstLicCounter", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumInstLicCounter")), SWInstCounter));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumFreeLicCounter", _.SUBT(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumInstLicCounter"))));

        }
        //----------------------------------------------------------------------------------------------------------
        public void CalcFolderLicCounter(ref object hlContext, ref object pDict)
        {
            object SWRefLicCounter = null;
            object SWInstCounter = null;
            object SoftwareLicense = null;
            object objType = null;
            object lstLicStatus = null;

            SWRefLicCounter = (Int16)0;
            SWInstCounter = (Int16)0;
            SoftwareLicense = VBScriptConstants.Nothing;
            objType = "";
            lstLicStatus = "";

            var enumerationContent4 = _.ENUMERABLE(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                SoftwareLicense = enumerationContent4.Current;
                objType = _.VAL(_.CALLm1argp(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.OR(_.EQ(_.NullableSTR(objType), "LicenseFolder"), _.EQ(_.NullableSTR(objType), "Software"))))
                {
                    object hlContext_vref24 = hlContext;
                    try
                    {
                        SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref24, v81 => { hlContext_vref24 = v81; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref24; }
                    _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumRefLicCounter", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));

                    object hlContext_vref25 = hlContext;
                    try
                    {
                        SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref25, v82 => { hlContext_vref25 = v82; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.InstalledLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref25; }
                    if (_.IF(_.GT(SWInstCounter, _.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumInstLicCounter")))))
                    {
                        _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumInstLicCounter", _.VAL(SWInstCounter));
                    }
                }
                if (_.IF(_.EQ(_.NullableSTR(objType), "SoftwareLicense")))
                {
                    lstLicStatus = _.VAL(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseDetail.LicenseStatus", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object hlContext_vref26 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref26, v83 => { hlContext_vref26 = v83; }).Val(_.CALLm1v5(this, _.NnO(SoftwareLicense, "SoftwareLicense"), "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref26; }
                        _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumRefLicCounter", _.ADD(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "SumFreeLicCounter", _.SUBT(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SumInstLicCounter"))));
        }
        //----------------------------------------------------------------------------------------------------------
        //Diese Function überprüft den ganzzahligen Wert (Integer).
        public object CheckIntegerValue(ref object hlContext, ref object intval)
        {
            object CheckIntegerValue_retVal = null;
            if (_.IF(_.OR(_.EQ(_.NullableSTR(intval), ""), _.EQ(_.ISNUMERIC(intval), false))))
            {
                CheckIntegerValue_retVal = (Int16)0;
            }
            else
            {
                CheckIntegerValue_retVal = _.CLNG(intval);
            }
            return CheckIntegerValue_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        public object OnCreate_HasAssociationToDelete(ref object hlContext, ref object AscDefName, ref object hlObjB)
        {
            object OnCreate_HasAssociationToDelete_retVal = null;
            object result = null;
            object cAssociationChanges = null;
            object oAssociationChange = null;
            object AscDefNameChange = null;
            object ixAC = null;
            result = false;
            cAssociationChanges = (Int16)0;
            cAssociationChanges = _.VAL(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart = _.NUM((Int16)0, loopEnd, (Int16)1);
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (ixAC = loopStart; _.StrictLTE(ixAC, loopEnd); ixAC = _.ADD(ixAC, (Int16)1))
                {
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v84 => { ixAC = v84; })));

                    AscDefNameChange = _.VAL(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "IsToDelete")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, _.NnO(hlObjB, "hlObjB"), "GetID"), _.CALLm2v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "EndB", "GetID"))))
                            {
                                result = true;
                                break;
                            } //check the ids
                        } // check the defnames
                    } // is to create
                }
            }
            OnCreate_HasAssociationToDelete_retVal = _.VAL(result);
            return OnCreate_HasAssociationToDelete_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        public object OnCreate_HasAssociationToCreate(ref object hlContext, ref object AscDefName, ref object hlObjB)
        {
            object OnCreate_HasAssociationToCreate_retVal = null;
            object result = null;
            object cAssociationChanges = null;
            object oAssociationChange = null;
            object AscDefNameChange = null;
            object ixAC = null;
            result = false;
            cAssociationChanges = (Int16)0;
            cAssociationChanges = _.VAL(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd2 = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart2 = _.NUM((Int16)0, loopEnd2, (Int16)1);
            if (_.StrictLTE(loopStart2, loopEnd2))
            {
                for (ixAC = loopStart2; _.StrictLTE(ixAC, loopEnd2); ixAC = _.ADD(ixAC, (Int16)1))
                {
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v85 => { ixAC = v85; })));

                    AscDefNameChange = _.VAL(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "IsToCreate")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, _.NnO(hlObjB, "hlObjB"), "GetID"), _.CALLm2v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "EndB", "GetID"))))
                            {
                                result = true;
                                break;
                            } //check the ids
                        } // check the defnames
                    } // is to create
                }
            }
            OnCreate_HasAssociationToCreate_retVal = _.VAL(result);
            return OnCreate_HasAssociationToCreate_retVal;
        }
        public object OnDelete_HasAssociationToCreate(ref object hlContext, ref object AscDefName, ref object hlObjB)
        {
            object OnDelete_HasAssociationToCreate_retVal = null;
            object result = null;
            object cAssociationChanges = null;
            object oAssociationChange = null;
            object AscDefNameChange = null;
            object ixAC = null;
            // bool
            result = false;

            //Anzahl der zu erstellenden oder löschenden Assoziationen
            cAssociationChanges = (Int16)0;
            cAssociationChanges = _.VAL(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd3 = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart3 = _.NUM((Int16)0, loopEnd3, (Int16)1);
            if (_.StrictLTE(loopStart3, loopEnd3))
            {
                for (ixAC = loopStart3; _.StrictLTE(ixAC, loopEnd3); ixAC = _.ADD(ixAC, (Int16)1))
                {

                    //Für jede Assoziations Änderung wird das entsprechende Infos (Objekt    ) ausgelsen.
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v86 => { ixAC = v86; })));
                    //Def Name der Assoc ermitteln, die angelegt werden soll
                    AscDefNameChange = _.VAL(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "IsToCreate")))
                    {
                        //Überprüfen ob die gewünschte Assoc auch angelegt werden soll.
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, _.NnO(hlObjB, "hlObjB"), "GetID"), _.CALLm2v0(this, _.NnO(oAssociationChange, "oAssociationChange"), "EndB", "GetID"))))
                            {
                                result = true;
                                break;
                            } //check the ids
                        } // check the defnames
                    } // is to create
                }
            }
            OnDelete_HasAssociationToCreate_retVal = _.VAL(result);
            return OnDelete_HasAssociationToCreate_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        public object GetAssociatedOrganizationalUnit(ref object hlContext, ref object lcid, ref object hlChild, ref object pDict, ref object outParentDefName)
        {
            object GetAssociatedOrganizationalUnit_retVal = null;
            object rsltParent = null;
            object objParent = null;
            GetAssociatedOrganizationalUnit_retVal = "";
            outParentDefName = "";

            rsltParent = "";
            object pDict_vref9 = pDict;
            try
            {
                rsltParent = _.VAL(_.CALLm1argp(this, _.NnO(hlChild, "hlChild"), "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).RefIfArray(pDict_vref9, _.ARGS.Val("AssocID"))));
            }
            finally { pDict = pDict_vref9; }
            if (_.IF(_.GTE(_.UBOUND(rsltParent), _.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ParentCounter")))))
            {
                objParent = VBScriptConstants.Nothing;
                var enumerationContent5 = _.ENUMERABLE(rsltParent).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent5.MoveNext())
                        break;
                    objParent = enumerationContent5.Current;
                    object pDict_vref10 = pDict;
                    try
                    {
                        GetAssociatedOrganizationalUnit_retVal = _.VAL(_.CALLm1argp(this, _.NnO(objParent, "objParent"), "GetValue", _.ARGS.RefIfArray(pDict_vref10, _.ARGS.Val("AttrName")).Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { pDict = pDict_vref10; }
                    object lcid_vref = lcid;
                    try
                    {
                        outParentDefName = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetDisplayName", _.ARGS.Val(_.CALLm1v5(this, _.NnO(objParent, "objParent"), "GetValue", "HLOBJECTINFO.DEFID", (Int16)0, (Int16)0, (Int16)0, (Int16)0)).Ref(lcid_vref, v87 => { lcid_vref = v87; })));
                    }
                    finally { lcid = lcid_vref; }
                    break;
                }
            }
            return GetAssociatedOrganizationalUnit_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------
        public object MIG_CreateXMLDocument(ref object hlSrvContext, ref object pDict)
        {
            object MIG_CreateXMLDocument_retVal = null;
            object objXMLDoc = null;
            object xmlProInc = null;
            object xmlRoot = null;
            object nodeSession = null;

            //XML-Objekt erstellen
            objXMLDoc = VBScriptConstants.Nothing;
            objXMLDoc = _.CREATEOBJECT("Msxml2.DOMDocument");

            //XML-Processing Instruction hinzufügen
            xmlProInc = VBScriptConstants.Nothing;
            xmlProInc = _.OBJ(_.CALLm1v2(this, _.NnO(objXMLDoc, "objXMLDoc"), "createProcessingInstruction", "xml", "version='1.0' encoding='UTF-8'"));
            _.CALLm1argp(this, _.NnO(objXMLDoc, "objXMLDoc"), "insertBefore", _.ARGS.Ref(xmlProInc, v88 => { xmlProInc = v88; }).Val(_.CALLm1v0(this, _.NnO(objXMLDoc, "objXMLDoc"), "firstChild")));

            //Root-Element erstellen
            xmlRoot = _.OBJ(_.CALLm1v1(this, _.NnO(objXMLDoc, "objXMLDoc"), "CreateElement", "ASAPBatch"));
            _.CALLm1v1(this, _.NnO(objXMLDoc, "objXMLDoc"), "AppendChild", xmlRoot);
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "xmlns", "http://www.brainware.ch/operationsmanager/asap-batch/1.1");
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "xmlns:dt", "http://www.brainware.ch/operationsmanager/wf/changemanagement/columbus/datatypes/1.1");
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "xsi:schemaLocation", "http://www.brainware.ch/operationsmanager/asap-batch/1.1 asap-batch-1.1.xsd");
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "version", "1.1");
            _.CALLm1v2(this, _.NnO(xmlRoot, "xmlRoot"), "SetAttribute", "responseRequired", "Yes");

            //Das Node Session hinzufügen
            nodeSession = _.OBJ(_.CALLm1v1(this, _.NnO(objXMLDoc, "objXMLDoc"), "CreateElement", "Session"));
            _.CALLm1v1(this, _.NnO(xmlRoot, "xmlRoot"), "AppendChild", nodeSession);
            _.CALLm1v2(this, _.NnO(nodeSession, "nodeSession"), "SetAttribute", "id", "s1");
            _.CALLm1v2(this, _.NnO(nodeSession, "nodeSession"), "SetAttribute", "loginname", "foreignSystems\\assetcolumbus");
            _.CALLm1v2(this, _.NnO(nodeSession, "nodeSession"), "SetAttribute", "password", "");

            //XML Dokument inkl. Header an das Dictionary übergeben
            _.SETm0a1(this, _.NnO(pDict, "pDict"), "XMLDocument", _.OBJ(objXMLDoc));
            return MIG_CreateXMLDocument_retVal;
        }
        //---------------------------------------------------------------------
        public object MIG_CreateADDXML2Columbus(ref object hlSrvContext, ref object pDict)
        {
            object MIG_CreateADDXML2Columbus_retVal = null;
            object xmlRoot = null;
            object nodeCreateInstanceRq = null;
            object nodeObserverKey = null;
            object nodeContextData = null;
            object nodeAddDeviceActualParams = null;
            object nodeDeviceIdentification = null;
            object nodeDeviceName = null;
            object nodeCmpyName = null;
            object nodeDomain = null;
            object nodeCostCenter = null;
            object nodeMACAddress = null;
            object nodeSubnetMask = null;
            object nodeHWType = null;
            object nodeOSType = null;
            object nodeActState = null;

            //Root Element aus dem XML ermitteln.
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, _.NnO(xmlRoot, "xmlRoot"), "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "id", "e7");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_adddevice");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeObserverKey);
            _.SETm1a0(this, _.NnO(nodeObserverKey, "nodeObserverKey"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ObserverKey"))));

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ContextData"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeAddDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:AddDeviceActualParams"));
            _.CALLm1v1(this, _.NnO(nodeContextData, "nodeContextData"), "AppendChild", nodeAddDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDeviceName);
            _.SETm1a0(this, _.NnO(nodeDeviceName, "nodeDeviceName"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("DeviceName"))));

            //Das Node CompanyName hinzufügen
            nodeCmpyName = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:CompanyName"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeCmpyName);
            _.SETm1a0(this, _.NnO(nodeCmpyName, "nodeCmpyName"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("CompanyName"))));

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDomain);
            _.SETm1a0(this, _.NnO(nodeDomain, "nodeDomain"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("Domain"))));

            //Das Node CostCenter hinzufügen
            nodeCostCenter = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:CostCenter"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeCostCenter);
            _.SETm1a0(this, _.NnO(nodeCostCenter, "nodeCostCenter"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("CostCenter"))));

            //Das Node MACAdess hinzufügen
            nodeMACAddress = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:MACAddress"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeMACAddress);
            _.SETm1a0(this, _.NnO(nodeMACAddress, "nodeMACAddress"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("MACAddress"))));

            //Das Node SubnetMask hinzufügen
            nodeSubnetMask = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:SubnetMask"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeSubnetMask);
            _.SETm1a0(this, _.NnO(nodeSubnetMask, "nodeSubnetMask"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SubnetMask"))));

            //Das Node HwTypeId hinzufügen
            nodeHWType = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:HwTypeId"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeHWType);
            _.SETm1a0(this, _.NnO(nodeHWType, "nodeHWType"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("HwTypeId"))));

            //Das Node OsTypeId hinzufügen
            nodeOSType = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:OsTypeId"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeOSType);
            _.SETm1a0(this, _.NnO(nodeOSType, "nodeOSType"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("OsTypeId"))));

            //Das Node ActivationState hinzufügen
            nodeActState = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:ActivationState"));
            _.CALLm1v1(this, _.NnO(nodeAddDeviceActualParams, "nodeAddDeviceActualParams"), "AppendChild", nodeActState);
            _.SETm1a0(this, _.NnO(nodeActState, "nodeActState"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ActivationState"))));

            return MIG_CreateADDXML2Columbus_retVal;
        }
        //---------------------------------------------------------------------
        public object MIG_CreateCHGXML2Columbus(ref object hlSrvContext, ref object pDict)
        {
            object MIG_CreateCHGXML2Columbus_retVal = null;
            object xmlRoot = null;
            object nodeCreateInstanceRq = null;
            object nodeObserverKey = null;
            object nodeContextData = null;
            object nodeChgDeviceActualParams = null;
            object nodeDeviceIdentification = null;
            object nodeDeviceName = null;
            object nodeDomain = null;
            object nodeCmpyName = null;
            object nodeCostCenter = null;
            object nodeMACAddress = null;
            object nodeSubnetMask = null;
            object nodeHWType = null;
            object nodeOSType = null;
            object nodeActState = null;

            //Root Element aus dem XML ermitteln.
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, _.NnO(xmlRoot, "xmlRoot"), "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "id", "e7");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_chgdevice");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeObserverKey);
            _.SETm1a0(this, _.NnO(nodeObserverKey, "nodeObserverKey"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ObserverKey"))));

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ContextData"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeChgDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:ChangeDeviceActualParams"));
            _.CALLm1v1(this, _.NnO(nodeContextData, "nodeContextData"), "AppendChild", nodeChgDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDeviceName);
            _.SETm1a0(this, _.NnO(nodeDeviceName, "nodeDeviceName"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("DeviceName"))));

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDomain);
            _.SETm1a0(this, _.NnO(nodeDomain, "nodeDomain"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("Domain"))));

            //Das Node CompanyName hinzufügen
            nodeCmpyName = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:CompanyName"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeCmpyName);
            _.SETm1a0(this, _.NnO(nodeCmpyName, "nodeCmpyName"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("CompanyName"))));

            //Das Node CostCenter hinzufügen
            nodeCostCenter = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:CostCenter"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeCostCenter);
            _.SETm1a0(this, _.NnO(nodeCostCenter, "nodeCostCenter"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("CostCenter"))));

            //Das Node MACAdess hinzufügen
            nodeMACAddress = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:MACAddress"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeMACAddress);
            _.SETm1a0(this, _.NnO(nodeMACAddress, "nodeMACAddress"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("MACAddress"))));

            //Das Node SubnetMask hinzufügen
            nodeSubnetMask = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:SubnetMask"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeSubnetMask);
            _.SETm1a0(this, _.NnO(nodeSubnetMask, "nodeSubnetMask"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("SubnetMask"))));

            //Das Node HwTypeId hinzufügen
            nodeHWType = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:HwTypeId"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeHWType);
            _.SETm1a0(this, _.NnO(nodeHWType, "nodeHWType"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("HwTypeId"))));

            //Das Node OsTypeId hinzufügen
            nodeOSType = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:OsTypeId"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeOSType);
            _.SETm1a0(this, _.NnO(nodeOSType, "nodeOSType"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("OsTypeId"))));

            //Das Node ActivationState hinzufügen
            nodeActState = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:ActivationState"));
            _.CALLm1v1(this, _.NnO(nodeChgDeviceActualParams, "nodeChgDeviceActualParams"), "AppendChild", nodeActState);
            _.SETm1a0(this, _.NnO(nodeActState, "nodeActState"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ActivationState"))));

            return MIG_CreateCHGXML2Columbus_retVal;
        }
        //---------------------------------------------------------------------
        public object MIG_CreateDELXML2Columbus(ref object hlSrvContext, ref object pDict)
        {
            object MIG_CreateDELXML2Columbus_retVal = null;
            object xmlRoot = null;
            object nodeCreateInstanceRq = null;
            object nodeObserverKey = null;
            object nodeContextData = null;
            object nodeRemoveDeviceActualParams = null;
            object nodeDeviceIdentification = null;
            object nodeDeviceName = null;
            object nodeDomain = null;

            //Root Element aus dem XML ermitteln.
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, _.NnO(xmlRoot, "xmlRoot"), "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "id", "e7");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_removedevice");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeObserverKey);
            _.SETm1a0(this, _.NnO(nodeObserverKey, "nodeObserverKey"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("ObserverKey"))));

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "ContextData"));
            _.CALLm1v1(this, _.NnO(nodeCreateInstanceRq, "nodeCreateInstanceRq"), "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeRemoveDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:RemoveDeviceActualParams"));
            _.CALLm1v1(this, _.NnO(nodeContextData, "nodeContextData"), "AppendChild", nodeRemoveDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, _.NnO(nodeRemoveDeviceActualParams, "nodeRemoveDeviceActualParams"), "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDeviceName);
            _.SETm1a0(this, _.NnO(nodeDeviceName, "nodeDeviceName"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("DeviceName"))));

            //Das Node CompanyName hinzufügen
            //Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
            //nodeDeviceIdentification.AppendChild (nodeCmpyName)
            //nodeCmpyName.Text = pDict("CompanyName")

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.NnO(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("XMLDocument")), "(_.call result)"), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, _.NnO(nodeDeviceIdentification, "nodeDeviceIdentification"), "AppendChild", nodeDomain);
            _.SETm1a0(this, _.NnO(nodeDomain, "nodeDomain"), "Text", _.VAL(_.CALLm0argp(this, _.NnO(pDict, "pDict"), _.ARGS.Val("Domain"))));

            return MIG_CreateDELXML2Columbus_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //Wenn beide Werte ein Datum sind, muss geprüft werden ob das Enddatum nach dem
        //Start Datum liegt. Falls nicht wird "False" zurückgegeben.
        public object MigCheckDatePeriod(ref object hlContext, ref object StartDate, ref object EndDate)
        {
            object MigCheckDatePeriod_retVal = null;
            MigCheckDatePeriod_retVal = false;

            if (_.IF(_.NOTEQ(_.NullableSTR(_.DATEPART("d", _.CDATE(StartDate))), "0")))
            {
                if (_.IF(_.LT(_.DATEPART("d", _.CDATE(StartDate)), _.DATEPART("d", _.CDATE(EndDate)))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    MigCheckDatePeriod_retVal = true;
                }

                if (_.IF(_.GT(_.DATEPART("yyyy", _.CDATE(StartDate)), _.DATEPART("yyyy", _.CDATE(EndDate)))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    if (_.IF(_.GT(_.DATEPART("y", _.CDATE(StartDate)), _.DATEPART("y", _.CDATE(EndDate)))))
                    {
                        if (_.IF(_.LT(_.DATEPART("yyyy", _.CDATE(StartDate)), _.DATEPART("yyyy", _.CDATE(EndDate)))))
                        {
                            MigCheckDatePeriod_retVal = true;
                        }
                        else
                        {
                            MigCheckDatePeriod_retVal = false;
                        }
                    }
                    else
                    {
                        MigCheckDatePeriod_retVal = true;
                    }
                }
            }
            return MigCheckDatePeriod_retVal;
        }
        //---------------------------------------------------------------------
        public object MIG_CheckCostCenter(ref object hlSrvContext, ref object strCostCenter)
        {
            object MIG_CheckCostCenter_retVal = null;
            object srchQuery = null;
            object Qry = null;
            object rsltQuery = null;
            MIG_CheckCostCenter_retVal = false;

            srchQuery = "";
            srchQuery = _.CONCAT("SEARCH Division WHERE OrganizationBilling.CostCenter_CA.CostCenter = \"", strCostCenter, "\"");
            Qry = VBScriptConstants.Nothing;
            rsltQuery = "";
            Qry = _.OBJ(_.CALLm1argp(this, _.NnO(hlSrvContext, "hlSrvContext"), "OpenSearch", _.ARGS.Ref(srchQuery, v89 => { srchQuery = v89; })));
            rsltQuery = _.VAL(_.CALLm1v4(this, _.NnO(Qry, "Qry"), "GetItems", (Int16)0, _.SUBT((Int16)1), _.SUBT((Int16)1), (Int16)0));
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltQuery)), (Int16)0)))
            {
                MIG_CheckCostCenter_retVal = true;
            }

            return MIG_CheckCostCenter_retVal;
        }
        public object CheckAgentHasMIGPartnerID(ref object hlContext, ref object relObjMIGPartnerID)
        {
            object CheckAgentHasMIGPartnerID_retVal = null;
            object flagAuthorized = null;
            object intAgentID = null;
            object objPerson = null;
            object strPersonInternalMIGPartnerIDs = null;
            //BOOL

            flagAuthorized = false;
            intAgentID = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetAgentID", _.ARGS.ForceBrackets()));
            objPerson = VBScriptConstants.Nothing;

            objPerson = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetPersonOfAgent", _.ARGS.Ref(intAgentID, v90 => { intAgentID = v90; })));

            bool ifResult6;
            object hlContext_vref27 = hlContext;
            try
            {
                ifResult6 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(hlContext_vref27, v93 => { hlContext_vref27 = v93; }).Ref(objPerson, v94 => { objPerson = v94; })), true));
            }
            finally { hlContext = hlContext_vref27; }
            if (ifResult6)
            {

                if (_.IF(_.NOTEQ(_.NullableSTR(relObjMIGPartnerID), "")))
                {

                    strPersonInternalMIGPartnerIDs = _.VAL(_.CALLm1v5(this, _.NnO(objPerson, "objPerson"), "GetValue", "MIGAgentInformation.InternalMIGPartnerID", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

                    if (_.IF(_.GT(_.NullableNUM(_.INSTR(strPersonInternalMIGPartnerIDs, relObjMIGPartnerID)), (Int16)0)))
                    {
                        flagAuthorized = true;
                    }
                }
                else
                {
                    //If relObjMIGPartnerID is Null or empty, modification allowed
                    flagAuthorized = true;
                }

            }

            //return
            CheckAgentHasMIGPartnerID_retVal = _.VAL(flagAuthorized);
            return CheckAgentHasMIGPartnerID_retVal;
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
