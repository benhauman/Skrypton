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
            //Globale Konstanten fuer freie Assoziationsdefinitionen

            //----------------------------------------------------------------------------------------------------------

            //----------------------------------------------------------------------------------------------------------
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
        //Wenn der Parameter "GetAll" auf False steht wird als Rueckgabewert fuer die Funktion
        //ebenfalls "False" ausgegben, wenn mehr als ein Standardeintrag gefunden wird.
        //Wenn fuer den Parameter "True" angeben wird, prueft die Funktion ob es tatsaechlich
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
                ItemIDs = _.VAL(_.CALLm1argp(this, hlObject, "GetContentIDs", _.ARGS.RefIfArray(dict_vref, _.ARGS.Val("Compound")).Val((Int16)0)));
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
                        strValue = _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.RefIfArray(dict_vref3, _.ARGS.Val("Value")).Val((Int16)0).Ref(Item, v4 => { Item = v4; }).Val((Int16)0).Val((Int16)0)));
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
                _.SETm0a1(this, dict, "DefValue", _.VAL(strValue));
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
                _.CALLm1argp(this, hlContext, "trace", _.ARGS.Val((Int16)1).Ref(text_vref, v5 => { text_vref = v5; }));
            }
            finally { text = text_vref; }
        }
        //---------------------------------------------------------------
        //Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
        public void SetPersonInformation(ref object hlContext, ref object hlObject, ref object dict)
        {
            object AttrDef = null;
            object strAttrValue = null;
            //Aus dem Dictionary wird das Attribut und der dazugehoerige Wert ermittelt.
            AttrDef = "";
            AttrDef = _.CONCAT("PersonInformation.", _.CALLm0argp(this, dict, _.ARGS.Val("PersInfoAttr")));

            strAttrValue = "";
            strAttrValue = _.VAL(_.CALLm0argp(this, dict, _.ARGS.Val("DefValue")));

            _.CALLm1argp(this, hlObject, "SetValue", _.ARGS.Ref(AttrDef, v6 => { AttrDef = v6; }).Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strAttrValue, v7 => { strAttrValue = v7; }));
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
            IsHLObject_retVal = _.VAL(_.AND(_.EQ(_.ISOBJECT(hlObject), true), _.EQ(_.IS(hlObject, VBScriptConstants.Nothing), false)));
            return IsHLObject_retVal;
        }
        //-------------------------------------------------------------------
        public object GetBaseType(ref object hlContext, ref object hlObject)
        {
            return _.VAL(_.CALLm1v5(this, hlObject, "GetValue", "HLOBJECTINFO.BASETYPE", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
        }
        //---------------------------------------------------------------
        //Dies ist eine rekursive Function zum ermitteln der Organisationshierarchie,
        //ausgehend vom der ersten OU ueberhalb einer Person.
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
                strOrgUnits = _.VAL(_.CALLm1v5(this, hlOrgUnit, "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            }
            else
            {
                strOrgUnits = _.CONCAT(strOrgUnits, ", ", _.CALLm1v5(this, hlOrgUnit, "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            }

            //Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //fuer die naechste Abfrage gewaehlt werden kann.
            orgaType = "";
            orgaType = _.VAL(_.CALLm1v0(this, hlOrgUnit, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Division")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, hlOrgUnit, "GetItems", 65536, (Int16)0, (Int16)0, "CompanyView"));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Site")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, hlOrgUnit, "GetItems", 65536, (Int16)0, (Int16)0, "Site2Company"));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Company")))
            {
                NextOrgUnit = _.VAL(_.CALLm1v4(this, hlOrgUnit, "GetItems", 65536, (Int16)0, (Int16)0, "Company2Company"));
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
        //ueber diese Function wird fuer ein Flag Attribut immer der Wert
        //True oder False ausgegeben.
        public object GetFlagValue(ref object hlContext, ref object hlObject, ref object hlattribute, ref object hlcontentid, ref object hlsuid)
        {
            object GetFlagValue_retVal = null;
            object hlattribute_vref = hlattribute, hlcontentid_vref = hlcontentid, hlsuid_vref = hlsuid;
            try
            {
                GetFlagValue_retVal = _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.Ref(hlattribute_vref, v11 => { hlattribute_vref = v11; }).Val((Int16)0).Ref(hlcontentid_vref, v12 => { hlcontentid_vref = v12; }).Ref(hlsuid_vref, v13 => { hlsuid_vref = v13; }).Val((Int16)0)));
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
        //Woerterbuch ohne Parameter.
        public object GetErrMsg0(ref object hlContext, ref object LocaleID, ref object ErrCode)
        {
            object GetErrMsg0_retVal = null;
            object strErrMsg = null;
            GetErrMsg0_retVal = "";

            strErrMsg = "";
            object ErrCode_vref = ErrCode, LocaleID_vref = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(ErrCode_vref, v14 => { ErrCode_vref = v14; }).Ref(LocaleID_vref, v15 => { LocaleID_vref = v15; })));
            }
            finally { ErrCode = ErrCode_vref; LocaleID = LocaleID_vref; }
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbNewLine, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
            //Rueckgabewert der Function ist die Fehlermeldung.
            GetErrMsg0_retVal = _.REPLACE(strErrMsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg0_retVal;
        }
        //Das Script ermittelt auf Basis der ersten uebergeordneten OU den gesamten Pfad bis zur Firma oder Konzern
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
            FirstOrgUnit = _.OBJ(_.CALLm1v0(this, hlContext, "GetRelatedObject"));

            bool ifResult;
            object hlContext_vref4 = hlContext;
            try
            {
                ifResult = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(hlContext_vref4, v18 => { hlContext_vref4 = v18; }).Ref(FirstOrgUnit, v19 => { FirstOrgUnit = v19; })), true));
            }
            finally { hlContext = hlContext_vref4; }
            if (ifResult)
            {
                if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.CALLm1v0(this, FirstOrgUnit, "GetType")), "Company"), _.NOTEQ(_.NullableSTR(_.CALLm1v0(this, FirstOrgUnit, "GetType")), "Division"))))
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
                rsltOrgUnit = _.VAL(_.CALLm1v4(this, hlPerson, "GetItems", 65536, (Int16)0, (Int16)0, "Person2Organization"));
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltOrgUnit)), (Int16)0)))
                {
                    FirstOrgUnit = _.OBJ(_.CALLm0argp(this, rsltOrgUnit, _.ARGS.Val((Int16)0)));
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

                    _.SETm0a1(this, dict, "DefValue", _.VAL(strOrgUnits));
                    _.SETm0a1(this, dict, "PersInfoAttr", "PersonOrganization");
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
        //Prozedur fuellt die Umzugshistorie fuer das entsprechende Objekt
        public void SetAssetHistory(ref object hlContext, ref object hlObjectA, ref object hlObjectB, ref object created)
        {
            object productDefName = null;
            object agentID = null;
            object contentID = null;
            object personOfAgent = null;
            object personName = null;
            object orgUnitName = null;
            object strErrMsg = null;

            productDefName = _.VAL(_.CALLm1argp(this, hlObjectB, "GetType", _.ARGS.ForceBrackets()));

            if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(productDefName), "Software"), _.NOTEQ(_.NullableSTR(productDefName), "SoftwareLicence"))))
            {
                contentID = _.VAL(_.CALLm1argp(this, hlObjectB, "GenerateContentID", _.ARGS.ForceBrackets()));
                agentID = _.VAL(_.CALLm1argp(this, hlContext, "GetAgentID", _.ARGS.ForceBrackets()));
                orgUnitName = _.VAL(_.CALLm1v5(this, hlObjectA, "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                personOfAgent = _.OBJ(_.CALLm1argp(this, hlContext, "GetPersonOfAgent", _.ARGS.Ref(agentID, v38 => { agentID = v38; })));
                if (_.IF(_.IS(personOfAgent, VBScriptConstants.Nothing)))
                {
                    object hlContext_vref10 = hlContext;
                    try
                    {
                        strErrMsg = _.VAL(_.CALLm1argp(this, _outer, "GetErrMsg0", _.ARGS.Ref(hlContext_vref10, v39 => { hlContext_vref10 = v39; }).Val(_.CALLm1v0(this, hlContext_vref10, "GetLocaleID")).Val("#ERR_SETASSETHISTORY")));
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
                    personName = _.VAL(_.CALLm1v5(this, personOfAgent, "GetValue", "PersonGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    personName = _.CONCAT(personName, ", ");
                    personName = _.CONCAT(personName, _.CALLm1v5(this, personOfAgent, "GetValue", "PersonGeneral.GivenName", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                }
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedBy").Val((Int16)0).Ref(contentID, v42 => { contentID = v42; }).Val((Int16)0).Ref(personName, v43 => { personName = v43; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID").Val((Int16)0).Ref(contentID, v44 => { contentID = v44; }).Val((Int16)0).Ref(agentID, v45 => { agentID = v45; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangeDate").Val((Int16)0).Ref(contentID, v46 => { contentID = v46; }).Val((Int16)0).Val(_.NOW()));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnit").Val((Int16)0).Ref(contentID, v47 => { contentID = v47; }).Val((Int16)0).Ref(orgUnitName, v48 => { orgUnitName = v48; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnitID").Val((Int16)0).Ref(contentID, v49 => { contentID = v49; }).Val((Int16)0).Val(_.CALLm1argp(this, hlObjectA, "GetID", _.ARGS.ForceBrackets())));

                if (_.IF(_.EQ(created, true)))
                {
                    _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v50 => { contentID = v50; }).Val((Int16)0).Val("HistoryActionCreated"));
                }
                else
                {
                    _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v51 => { contentID = v51; }).Val((Int16)0).Val("HistoryActionDeleted"));
                }
            }
        }
        //---------------------------------------------------------------
        //Diese Function ermitellt eine Fehlermeldung aus dem helpLine
        //Woerterbuch mit einem Parameter.
        public object GetErrMsg1(ref object hlContext, ref object LocaleID, ref object ErrCode, ref object Arg1)
        {
            object GetErrMsg1_retVal = null;
            object strErrMsg = null;
            GetErrMsg1_retVal = "";

            strErrMsg = "";
            object ErrCode_vref2 = ErrCode, LocaleID_vref2 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(ErrCode_vref2, v52 => { ErrCode_vref2 = v52; }).Ref(LocaleID_vref2, v53 => { LocaleID_vref2 = v53; })));
            }
            finally { ErrCode = ErrCode_vref2; LocaleID = LocaleID_vref2; }
            strErrMsg = _.REPLACE(strErrMsg, "%1", Arg1);
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbLf, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
            //Rueckgabewert der Function ist die Fehlermeldung.
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
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(ErrCode_vref3, v54 => { ErrCode_vref3 = v54; }).Ref(LocaleID_vref3, v55 => { LocaleID_vref3 = v55; })));
            }
            finally { ErrCode = ErrCode_vref3; LocaleID = LocaleID_vref3; }
            strErrMsg = _.REPLACE(strErrMsg, "%1", Arg1);
            strErrMsg = _.REPLACE(strErrMsg, "%2", Arg2);
            strErrMsg = _.CONCAT(strErrMsg, VBScriptConstants.vbLf, "(Code: ", ErrCode, ")");

            //Den Paramenter %LF% durch Zeilenumbrueche ersetzen.
            //Rueckgabewert der Function ist die Fehlermeldung.
            GetErrMsg2_retVal = _.REPLACE(strErrMsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg2_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //In dieser Funktion wird geprueft, ob es unterhalb einer Software Suite
        //bereits Lizenzumschlaege mit Lizenzen gibt.
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

            //Pruefen ob es Software Lizenzobjekte/Lizenzumschlaege unterhalb des Folders gibt.
            object HLASC_SoftwareLicenseFolderView_vref = HLASC_SoftwareLicenseFolderView;
            try
            {
                rsltSWFolders = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(HLASC_SoftwareLicenseFolderView_vref, v56 => { HLASC_SoftwareLicenseFolderView_vref = v56; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = HLASC_SoftwareLicenseFolderView_vref; }

            var enumerationContent2 = _.ENUMERABLE(rsltSWFolders).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                SoftwareLicense = enumerationContent2.Current;
                objType = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objType), "LicenseFolder")))
                {
                    object hlContext_vref12 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref12, v57 => { hlContext_vref12 = v57; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref12; }
                    if (_.IF(_.GT(_.NullableNUM(GetReferenceLicenseCount_retVal), (Int16)0)))
                    {
                        return GetReferenceLicenseCount_retVal;
                    }
                }
                if (_.IF(_.AND(_.EQ(_.NullableSTR(objType), "SoftwareLicense"), _.EQ(_.CBOOL(chkFolderOnly), false))))
                {
                    object hlContext_vref13 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref13, v58 => { hlContext_vref13 = v58; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
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
        //den obersten Lizenz Umschlag ermittelt. Auf dem Weg dort hin wird geprueft ob einer
        //der Lizenzumschlaege eine Software Suite ist.
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
            //Start Folders existiert. Die Variable muss von aussen mit einem Startwert
            //initialisiert werden.
            if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareSuiteFolderLevel"))), (Int16)0), _.EQ(_.NullableSTR(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareSuiteFolderLevel"))), ""))))
            {
                _.SETm0a1(this, pDict, "SoftwareSuiteFolderLevel", (Int16)1);
            }
            else
            {
                _.SETm0a1(this, pDict, "SoftwareSuiteFolderLevel", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareSuiteFolderLevel")), (Int16)1));
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
                _.SETm0a1(this, pDict, "SoftwareSuiteFolder", _.VAL(_.CALLm1v5(this, hlParentSWFolder, "GetValue", "OrganizationGeneral.Name", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
                return CheckForSoftwareSuiteFolder_retVal;
            }

            //Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            object HLASC_SoftwareLicenseFolderView_vref2 = HLASC_SoftwareLicenseFolderView;
            try
            {
                NextSWFolder = _.VAL(_.CALLm1argp(this, hlParentSWFolder, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(HLASC_SoftwareLicenseFolderView_vref2, v61 => { HLASC_SoftwareLicenseFolderView_vref2 = v61; })));
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

            //Dictionary Eintraege initalisieren
            _.SETm0a1(this, pDict, "SoftwareLicenses", "");
            _.SETm0a1(this, pDict, "SumRefLicCounter", (Int16)0);
            _.SETm0a1(this, pDict, "SumInstLicCounter", (Int16)0);
            _.SETm0a1(this, pDict, "SumFreeLicCounter", (Int16)0);

            //Pruefen ob es Software Lizenzobjekte unterhalb des Folders gibt.
            object assocName_vref = assocName;
            try
            {
                _.SETm0a1(this, pDict, "SoftwareLicenses", _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(assocName_vref, v65 => { assocName_vref = v65; }))));
            }
            finally { assocName = assocName_vref; }

            //Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
            //Objekte gezaehlt werden muessen
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
                ifResult5 = _.IF(_.GTE(_.NullableNUM(_.UBOUND(_.CALLm0argp(this, pDict_vref2, _.ARGS.Val("SoftwareLicenses")))), (Int16)0));
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
            //Gesatmzahl der Lizenzen in den Lizenzumschlag zurueckschreiben
            object pDict_vref5 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref5, _.ARGS.Val("SumRefLicCounter")));
            }
            finally { pDict = pDict_vref5; }
            object pDict_vref6 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref6, _.ARGS.Val("SumInstLicCounter")));
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
                _.SETm0a1(this, pDict, "SumFreeLicCounter", (Int16)0);
            }
            object pDict_vref7 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.FreeLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref7, _.ARGS.Val("SumFreeLicCounter")));
            }
            finally { pDict = pDict_vref7; }

            //Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //fuer die naechste Abfrage gewaehlt werden kann.
            NextSWFolder = "";
            a = "";
            a = _.VAL(_.CALLm1v0(this, hlSWFolder, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(a), "LicenseFolder")))
            {
                object assocName_vref2 = assocName;
                try
                {
                    NextSWFolder = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Ref(assocName_vref2, v74 => { assocName_vref2 = v74; })));
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
            return _.VAL(_.AND(_.ISOBJECT(obj), _.NOT(_.IS(obj, VBScriptConstants.Nothing))));
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

            var enumerationContent3 = _.ENUMERABLE(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                SoftwareLicense = enumerationContent3.Current;
                objType = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objType), "SoftwareLicense")))
                {
                    lstLicStatus = _.VAL(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseDetail.LicenseStatus", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object hlContext_vref21 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref21, v78 => { hlContext_vref21 = v78; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref21; }
                        _.SETm0a1(this, pDict, "SumRefLicCounter", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                    }
                }
                else
                {
                    if (_.IF(_.OR(_.EQ(_.NullableSTR(objType), "LicenseFolder"), _.EQ(_.NullableSTR(objType), "Software"))))
                    {
                        object hlContext_vref22 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref22, v79 => { hlContext_vref22 = v79; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref22; }
                        _.SETm0a1(this, pDict, "SumRefLicCounter", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                        object hlContext_vref23 = hlContext;
                        try
                        {
                            SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref23, v80 => { hlContext_vref23 = v80; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.InstalledLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref23; }
                        _.SETm0a1(this, pDict, "SumInstLicCounter", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter")), SWInstCounter));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SETm0a1(this, pDict, "SumFreeLicCounter", _.SUBT(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter"))));

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

            var enumerationContent4 = _.ENUMERABLE(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                SoftwareLicense = enumerationContent4.Current;
                objType = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.OR(_.EQ(_.NullableSTR(objType), "LicenseFolder"), _.EQ(_.NullableSTR(objType), "Software"))))
                {
                    object hlContext_vref24 = hlContext;
                    try
                    {
                        SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref24, v81 => { hlContext_vref24 = v81; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref24; }
                    _.SETm0a1(this, pDict, "SumRefLicCounter", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));

                    object hlContext_vref25 = hlContext;
                    try
                    {
                        SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref25, v82 => { hlContext_vref25 = v82; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.InstalledLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                    }
                    finally { hlContext = hlContext_vref25; }
                    if (_.IF(_.GT(SWInstCounter, _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter")))))
                    {
                        _.SETm0a1(this, pDict, "SumInstLicCounter", _.VAL(SWInstCounter));
                    }
                }
                if (_.IF(_.EQ(_.NullableSTR(objType), "SoftwareLicense")))
                {
                    lstLicStatus = _.VAL(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseDetail.LicenseStatus", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object hlContext_vref26 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(hlContext_vref26, v83 => { hlContext_vref26 = v83; }).Val(_.CALLm1v5(this, SoftwareLicense, "GetValue", "SoftwareLicenseCounter.ReferenceLicenseCount", (Int16)0, (Int16)0, (Int16)0, (Int16)0))));
                        }
                        finally { hlContext = hlContext_vref26; }
                        _.SETm0a1(this, pDict, "SumRefLicCounter", _.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SETm0a1(this, pDict, "SumFreeLicCounter", _.SUBT(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter"))));
        }
        //----------------------------------------------------------------------------------------------------------
        //Diese Function ueberprueft den ganzzahligen Wert (Integer).
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
            cAssociationChanges = _.VAL(_.CALLm1v0(this, hlContext, "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart = _.NUM((Int16)0, loopEnd, (Int16)1);
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (ixAC = loopStart; _.StrictLTE(ixAC, loopEnd); ixAC = _.ADD(ixAC, (Int16)1))
                {
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v84 => { ixAC = v84; })));

                    AscDefNameChange = _.VAL(_.CALLm1v0(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, oAssociationChange, "IsToDelete")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, hlObjB, "GetID"), _.CALLm2v0(this, oAssociationChange, "EndB", "GetID"))))
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
            cAssociationChanges = _.VAL(_.CALLm1v0(this, hlContext, "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd2 = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart2 = _.NUM((Int16)0, loopEnd2, (Int16)1);
            if (_.StrictLTE(loopStart2, loopEnd2))
            {
                for (ixAC = loopStart2; _.StrictLTE(ixAC, loopEnd2); ixAC = _.ADD(ixAC, (Int16)1))
                {
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v85 => { ixAC = v85; })));

                    AscDefNameChange = _.VAL(_.CALLm1v0(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, oAssociationChange, "IsToCreate")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, hlObjB, "GetID"), _.CALLm2v0(this, oAssociationChange, "EndB", "GetID"))))
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

            //Anzahl der zu erstellenden oder loeschenden Assoziationen
            cAssociationChanges = (Int16)0;
            cAssociationChanges = _.VAL(_.CALLm1v0(this, hlContext, "GetAssociationChangesCount"));

            oAssociationChange = VBScriptConstants.Nothing;
            AscDefNameChange = "";
            ixAC = (Int16)0;

            var loopEnd3 = _.NUM(_.SUBT(cAssociationChanges, (Int16)1));
            var loopStart3 = _.NUM((Int16)0, loopEnd3, (Int16)1);
            if (_.StrictLTE(loopStart3, loopEnd3))
            {
                for (ixAC = loopStart3; _.StrictLTE(ixAC, loopEnd3); ixAC = _.ADD(ixAC, (Int16)1))
                {

                    //Fuer jede Assoziations aenderung wird das entsprechende Infos (Objekt    ) ausgelsen.
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v86 => { ixAC = v86; })));
                    //Def Name der Assoc ermitteln, die angelegt werden soll
                    AscDefNameChange = _.VAL(_.CALLm1v0(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, oAssociationChange, "IsToCreate")))
                    {
                        //ueberpruefen ob die gewuenschte Assoc auch angelegt werden soll.
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v0(this, hlObjB, "GetID"), _.CALLm2v0(this, oAssociationChange, "EndB", "GetID"))))
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
                rsltParent = _.VAL(_.CALLm1argp(this, hlChild, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).RefIfArray(pDict_vref9, _.ARGS.Val("AssocID"))));
            }
            finally { pDict = pDict_vref9; }
            if (_.IF(_.GTE(_.UBOUND(rsltParent), _.CALLm0argp(this, pDict, _.ARGS.Val("ParentCounter")))))
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
                        GetAssociatedOrganizationalUnit_retVal = _.VAL(_.CALLm1argp(this, objParent, "GetValue", _.ARGS.Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(pDict_vref10, _.ARGS.Val("AttrName")).Val((Int16)0)));
                    }
                    finally { pDict = pDict_vref10; }
                    object lcid_vref = lcid;
                    try
                    {
                        outParentDefName = _.VAL(_.CALLm1argp(this, hlContext, "GetDisplayName", _.ARGS.Val(_.CALLm1v5(this, objParent, "GetValue", (Int16)0, (Int16)0, (Int16)0, (Int16)0, "HLOBJECTINFO.DEFID")).Ref(lcid_vref, v87 => { lcid_vref = v87; })));
                    }
                    finally { lcid = lcid_vref; }
                    break;
                }
            }
            return GetAssociatedOrganizationalUnit_retVal;
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
