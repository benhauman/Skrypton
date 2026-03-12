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

            _outer.HLASC_Software2Computer = "Software2Computer";
            _outer.HLASC_SoftwareLicenseGroupView = "LicenseGroupView";
            _outer.HLASC_SoftwareLicenseFolderView = "LicenseFolderView";
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
            HLASC_SoftwareLicenseFolderView = null;
            HLASC_SoftwareLicenseGroupView = null;
            HLASC_Software2Computer = null;
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
            object byrefalias = dict;
            try
            {
                ItemIDs = _.VAL(_.CALLm1argp(this, hlObject, "GetContentIDs", _.ARGS.RefIfArray(byrefalias, _.ARGS.Val("Compound")).Val((Int16)0)));
            }
            finally { dict = byrefalias; }

            Item = (Int16)0;
            var enumerationContent = _.ENUMERABLE(ItemIDs).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                Item = enumerationContent.Current;
                defItem = false;
                object byrefalias2 = hlContext, byrefalias3 = hlObject, byrefalias4 = dict;
                try
                {
                    defItem = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias2, v => { byrefalias2 = v; }).Ref(byrefalias3, v2 => { byrefalias3 = v2; }).RefIfArray(byrefalias4, _.ARGS.Val("Default")).Ref(Item, v3 => { Item = v3; }).Val((Int16)0)));
                }
                finally { hlContext = byrefalias2; hlObject = byrefalias3; dict = byrefalias4; }
                if (_.IF(_.EQ(_.CBOOL(defItem), true)))
                {
                    ItemCount = _.ADD(ItemCount, (Int16)1);
                    object byrefalias5 = dict;
                    try
                    {
                        strValue = _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.RefIfArray(byrefalias5, _.ARGS.Val("Value")).Val((Int16)0).Ref(Item, v4 => { Item = v4; }).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { dict = byrefalias5; }
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
                _.SET(_.VAL(strValue), this, dict, null, _.ARGS.Val("DefValue"));
            }
            return GetCommunicationDefault_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
        public void Trace(ref object hlContext, ref object text)
        {
            object byrefalias6 = text;
            try
            {
                _.CALLm1argp(this, hlContext, "trace", _.ARGS.Val((Int16)1).Ref(byrefalias6, v5 => { byrefalias6 = v5; }));
            }
            finally { text = byrefalias6; }
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
            object byrefalias7 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(byrefalias7, v8 => { byrefalias7 = v8; }).Val(_.CONCAT("Type ", _.VARTYPE(hlObject))));
            }
            finally { hlContext = byrefalias7; }
            IsHLObject_retVal = _.VAL(_.AND(_.EQ(_.ISOBJECT(hlObject), true), _.EQ(_.IS(hlObject, VBScriptConstants.Nothing), false)));
            return IsHLObject_retVal;
        }

        //-------------------------------------------------------------------
        public object GetBaseType(ref object hlContext, ref object hlObject)
        {
            return _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.Val("HLOBJECTINFO.BASETYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
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
                strOrgUnits = _.VAL(_.CALLm1argp(this, hlOrgUnit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            else
            {
                strOrgUnits = _.CONCAT(strOrgUnits, ", ", _.CALLm1argp(this, hlOrgUnit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            //Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //fuer die naechste Abfrage gewaehlt werden kann.
            orgaType = "";
            orgaType = _.VAL(_.CALLm1v(this, hlOrgUnit, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Division")))
            {
                NextOrgUnit = _.VAL(_.CALLm1argp(this, hlOrgUnit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("CompanyView")));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Site")))
            {
                NextOrgUnit = _.VAL(_.CALLm1argp(this, hlOrgUnit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Site2Company")));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgaType), "Company")))
            {
                NextOrgUnit = _.VAL(_.CALLm1argp(this, hlOrgUnit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Company2Company")));
            }

            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.ISARRAY(NextOrgUnit)))
            {
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextOrgUnit)), (Int16)0)))
                {
                    object byrefalias8 = hlContext, byrefalias9 = strOrgUnits;
                    try
                    {
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias8, v9 => { byrefalias8 = v9; }).RefIfArray(NextOrgUnit, _.ARGS.Val((Int16)0)).Ref(byrefalias9, v10 => { byrefalias9 = v10; })));
                    }
                    finally { hlContext = byrefalias8; strOrgUnits = byrefalias9; }
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
            object byrefalias10 = hlattribute, byrefalias11 = hlcontentid, byrefalias12 = hlsuid;
            try
            {
                GetFlagValue_retVal = _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.Ref(byrefalias10, v11 => { byrefalias10 = v11; }).Val((Int16)0).Ref(byrefalias11, v12 => { byrefalias11 = v12; }).Ref(byrefalias12, v13 => { byrefalias12 = v13; }).Val((Int16)0)));
            }
            finally { hlattribute = byrefalias10; hlcontentid = byrefalias11; hlsuid = byrefalias12; }
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
            object byrefalias13 = ErrCode, byrefalias14 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias13, v14 => { byrefalias13 = v14; }).Ref(byrefalias14, v15 => { byrefalias14 = v15; })));
            }
            finally { ErrCode = byrefalias13; LocaleID = byrefalias14; }
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
            FirstOrgUnit = _.OBJ(_.CALLm1v(this, hlContext, "GetRelatedObject"));

            bool ifResult;
            object byrefalias15 = hlContext;
            try
            {
                ifResult = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias15, v18 => { byrefalias15 = v18; }).Ref(FirstOrgUnit, v19 => { FirstOrgUnit = v19; })), true));
            }
            finally { hlContext = byrefalias15; }
            if (ifResult)
            {
                if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.CALLm1v(this, FirstOrgUnit, "GetType")), "Company"), _.NOTEQ(_.NullableSTR(_.CALLm1v(this, FirstOrgUnit, "GetType")), "Division"))))
                {
                    FirstOrgUnit = VBScriptConstants.Nothing;
                }
            }

            bool ifResult2;
            object byrefalias16 = hlContext;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias16, v22 => { byrefalias16 = v22; }).Ref(FirstOrgUnit, v23 => { FirstOrgUnit = v23; })), false));
            }
            finally { hlContext = byrefalias16; }
            if (ifResult2)
            {
                rsltOrgUnit = "";
                rsltOrgUnit = _.VAL(_.CALLm1argp(this, hlPerson, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Person2Organization")));
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltOrgUnit)), (Int16)0)))
                {
                    FirstOrgUnit = _.OBJ(_.CALLm0argp(this, rsltOrgUnit, _.ARGS.Val((Int16)0)));
                }
            }

            bool ifResult3;
            object byrefalias17 = hlContext;
            try
            {
                ifResult3 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias17, v26 => { byrefalias17 = v26; }).Ref(FirstOrgUnit, v27 => { FirstOrgUnit = v27; })), true));
            }
            finally { hlContext = byrefalias17; }
            if (ifResult3)
            {
                bool ifResult4;
                object byrefalias18 = hlContext;
                try
                {
                    ifResult4 = _.IF(_.EQ(_.NullableSTR(_.CALLm1argp(this, _outer, "GetBaseType", _.ARGS.Ref(byrefalias18, v30 => { byrefalias18 = v30; }).Ref(FirstOrgUnit, v31 => { FirstOrgUnit = v31; }))), "ORGANISATION"));
                }
                finally { hlContext = byrefalias18; }
                if (ifResult4)
                {
                    retval = "";
                    strOrgUnits = "";
                    object byrefalias19 = hlContext;
                    try
                    {
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias19, v32 => { byrefalias19 = v32; }).Ref(FirstOrgUnit, v33 => { FirstOrgUnit = v33; }).Ref(strOrgUnits, v34 => { strOrgUnits = v34; })));
                    }
                    finally { hlContext = byrefalias19; }

                    _.SET(_.VAL(strOrgUnits), this, dict, null, _.ARGS.Val("DefValue"));
                    _.SET("PersonOrganization", this, dict, null, _.ARGS.Val("PersInfoAttr"));
                    object byrefalias20 = hlContext, byrefalias21 = hlPerson, byrefalias22 = dict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "SetPersonInformation", _.ARGS.Ref(byrefalias20, v35 => { byrefalias20 = v35; }).Ref(byrefalias21, v36 => { byrefalias21 = v36; }).Ref(byrefalias22, v37 => { byrefalias22 = v37; }));
                    }
                    finally { hlContext = byrefalias20; hlPerson = byrefalias21; dict = byrefalias22; }
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
                orgUnitName = _.VAL(_.CALLm1argp(this, hlObjectA, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                personOfAgent = _.OBJ(_.CALLm1argp(this, hlContext, "GetPersonOfAgent", _.ARGS.Ref(agentID, v38 => { agentID = v38; })));
                if (_.IF(_.IS(personOfAgent, VBScriptConstants.Nothing)))
                {
                    object byrefalias23 = hlContext;
                    try
                    {
                        strErrMsg = _.VAL(_.CALLm1argp(this, _outer, "GetErrMsg0", _.ARGS.Ref(byrefalias23, v39 => { byrefalias23 = v39; }).Val(_.CALLm1v(this, byrefalias23, "GetLocaleID")).Val("#ERR_SETASSETHISTORY")));
                    }
                    finally { hlContext = byrefalias23; }
                    object byrefalias24 = hlContext;
                    try
                    {
                        _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(byrefalias24, v40 => { byrefalias24 = v40; }).Ref(strErrMsg, v41 => { strErrMsg = v41; }));
                    }
                    finally { hlContext = byrefalias24; }
                    //hlContext.abortcommand strErrMsg
                }
                else
                {
                    personName = _.VAL(_.CALLm1argp(this, personOfAgent, "GetValue", _.ARGS.Val("PersonGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    personName = _.CONCAT(personName, ", ");
                    personName = _.CONCAT(personName, _.CALLm1argp(this, personOfAgent, "GetValue", _.ARGS.Val("PersonGeneral.GivenName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
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
            object byrefalias25 = ErrCode, byrefalias26 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias25, v52 => { byrefalias25 = v52; }).Ref(byrefalias26, v53 => { byrefalias26 = v53; })));
            }
            finally { ErrCode = byrefalias25; LocaleID = byrefalias26; }
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
            object byrefalias27 = ErrCode, byrefalias28 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias27, v54 => { byrefalias27 = v54; }).Ref(byrefalias28, v55 => { byrefalias28 = v55; })));
            }
            finally { ErrCode = byrefalias27; LocaleID = byrefalias28; }
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
            object byrefalias29 = HLASC_SoftwareLicenseFolderView;
            try
            {
                rsltSWFolders = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias29, v56 => { byrefalias29 = v56; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = byrefalias29; }

            var enumerationContent2 = _.ENUMERABLE(rsltSWFolders).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                SoftwareLicense = enumerationContent2.Current;
                objType = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objType), "LicenseFolder")))
                {
                    object byrefalias30 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias30, v57 => { byrefalias30 = v57; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlContext = byrefalias30; }
                    if (_.IF(_.GT(_.NullableNUM(GetReferenceLicenseCount_retVal), (Int16)0)))
                    {
                        return GetReferenceLicenseCount_retVal;
                    }
                }
                if (_.IF(_.AND(_.EQ(_.NullableSTR(objType), "SoftwareLicense"), _.EQ(_.CBOOL(chkFolderOnly), false))))
                {
                    object byrefalias31 = hlContext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias31, v58 => { byrefalias31 = v58; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlContext = byrefalias31; }
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
                _.SET((Int16)1, this, pDict, null, _.ARGS.Val("SoftwareSuiteFolderLevel"));
            }
            else
            {
                _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SoftwareSuiteFolderLevel")), (Int16)1), this, pDict, null, _.ARGS.Val("SoftwareSuiteFolderLevel"));
            }

            //Amhand des Flags "Software Suite" festellen ob ein Lizenzumschlag als Software Suite
            //gekennzeichnet ist. Falls Ja, Name des Umschlags auslesen und Funktion abbrechen.
            CheckSoftwareSuite = false;
            object byrefalias32 = hlContext, byrefalias33 = hlParentSWFolder;
            try
            {
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias32, v59 => { byrefalias32 = v59; }).Ref(byrefalias33, v60 => { byrefalias33 = v60; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = byrefalias32; hlParentSWFolder = byrefalias33; }
            if (_.IF(_.EQ(_.CBOOL(CheckSoftwareSuite), true)))
            {
                _.SET(_.VAL(_.CALLm1argp(this, hlParentSWFolder, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))), this, pDict, null, _.ARGS.Val("SoftwareSuiteFolder"));
                return CheckForSoftwareSuiteFolder_retVal;
            }

            //Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            object byrefalias34 = HLASC_SoftwareLicenseFolderView;
            try
            {
                NextSWFolder = _.VAL(_.CALLm1argp(this, hlParentSWFolder, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias34, v61 => { byrefalias34 = v61; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = byrefalias34; }
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextSWFolder)), (Int16)0)))
            {
                object byrefalias35 = hlContext, byrefalias36 = pDict, byrefalias37 = HLASC_SoftwareLicenseFolderView;
                try
                {
                    retval = _.VAL(_.CALLm1argp(this, _outer, "CheckForSoftwareSuiteFolder", _.ARGS.Ref(byrefalias35, v62 => { byrefalias35 = v62; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(byrefalias36, v63 => { byrefalias36 = v63; }).Ref(byrefalias37, v64 => { byrefalias37 = v64; })));
                }
                finally { hlContext = byrefalias35; pDict = byrefalias36; HLASC_SoftwareLicenseFolderView = byrefalias37; }
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
            _.SET("", this, pDict, null, _.ARGS.Val("SoftwareLicenses"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumInstLicCounter"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumFreeLicCounter"));

            //Pruefen ob es Software Lizenzobjekte unterhalb des Folders gibt.
            object byrefalias38 = assocName;
            try
            {
                _.SET(_.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias38, v65 => { byrefalias38 = v65; }))), this, pDict, null, _.ARGS.Val("SoftwareLicenses"));
            }
            finally { assocName = byrefalias38; }

            //Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
            //Objekte gezaehlt werden muessen
            CheckSoftwareSuite = false;
            object byrefalias39 = hlContext, byrefalias40 = hlSWFolder;
            try
            {
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias39, v66 => { byrefalias39 = v66; }).Ref(byrefalias40, v67 => { byrefalias40 = v67; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = byrefalias39; hlSWFolder = byrefalias40; }

            bool ifResult5;
            object byrefalias41 = pDict;
            try
            {
                ifResult5 = _.IF(_.GTE(_.NullableNUM(_.UBOUND(_.CALLm0argp(this, byrefalias41, _.ARGS.Val("SoftwareLicenses")))), (Int16)0));
            }
            finally { pDict = byrefalias41; }
            if (ifResult5)
            {
                if (_.IF(_.EQ(_.CBOOL(CheckSoftwareSuite), false)))
                {
                    object byrefalias42 = hlContext, byrefalias43 = pDict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "CalcAllLicCounter", _.ARGS.Ref(byrefalias42, v68 => { byrefalias42 = v68; }).Ref(byrefalias43, v69 => { byrefalias43 = v69; }));
                    }
                    finally { hlContext = byrefalias42; pDict = byrefalias43; }
                }
                else
                {
                    object byrefalias44 = hlContext, byrefalias45 = pDict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "CalcFolderLicCounter", _.ARGS.Ref(byrefalias44, v70 => { byrefalias44 = v70; }).Ref(byrefalias45, v71 => { byrefalias45 = v71; }));
                    }
                    finally { hlContext = byrefalias44; pDict = byrefalias45; }
                }
            }
            //Gesatmzahl der Lizenzen in den Lizenzumschlag zurueckschreiben
            object byrefalias46 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias46, _.ARGS.Val("SumRefLicCounter")));
            }
            finally { pDict = byrefalias46; }
            object byrefalias47 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias47, _.ARGS.Val("SumInstLicCounter")));
            }
            finally { pDict = byrefalias47; }

            //Wenn die Lizenzkontrolle durch den Applikations Server erfolgt ("Lizenzkontrolle durch Server")
            //dann die Anzahl freier Lizenzen immer auf den Wert "0" setzen.
            CheckLicContrByServer = false;
            object byrefalias48 = hlContext, byrefalias49 = hlSWFolder;
            try
            {
                CheckLicContrByServer = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias48, v72 => { byrefalias48 = v72; }).Ref(byrefalias49, v73 => { byrefalias49 = v73; }).Val("SoftwareLicenseFolderDetail.FlagLicenseControlledByServer").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlContext = byrefalias48; hlSWFolder = byrefalias49; }
            if (_.IF(_.EQ(_.CBOOL(CheckLicContrByServer), true)))
            {
                _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumFreeLicCounter"));
            }
            object byrefalias50 = pDict;
            try
            {
                _.CALLm1argp(this, hlSWFolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.FreeLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias50, _.ARGS.Val("SumFreeLicCounter")));
            }
            finally { pDict = byrefalias50; }

            //Erst pruefen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //fuer die naechste Abfrage gewaehlt werden kann.
            NextSWFolder = "";
            a = "";
            a = _.VAL(_.CALLm1v(this, hlSWFolder, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(a), "LicenseFolder")))
            {
                object byrefalias51 = assocName;
                try
                {
                    NextSWFolder = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Ref(byrefalias51, v74 => { byrefalias51 = v74; })));
                }
                finally { assocName = byrefalias51; }
            }
            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextSWFolder)), (Int16)0)))
            {
                object byrefalias52 = hlContext, byrefalias53 = pDict, byrefalias54 = assocName;
                try
                {
                    retval = _.VAL(_.CALLm1argp(this, _outer, "SetLicenseCounter", _.ARGS.Ref(byrefalias52, v75 => { byrefalias52 = v75; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(byrefalias53, v76 => { byrefalias53 = v76; }).Ref(byrefalias54, v77 => { byrefalias54 = v77; })));
                }
                finally { hlContext = byrefalias52; pDict = byrefalias53; assocName = byrefalias54; }
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
                    lstLicStatus = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseDetail.LicenseStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object byrefalias55 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias55, v78 => { byrefalias55 = v78; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlContext = byrefalias55; }
                        _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
                    }
                }
                else
                {
                    if (_.IF(_.OR(_.EQ(_.NullableSTR(objType), "LicenseFolder"), _.EQ(_.NullableSTR(objType), "Software"))))
                    {
                        object byrefalias56 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias56, v79 => { byrefalias56 = v79; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlContext = byrefalias56; }
                        _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
                        object byrefalias57 = hlContext;
                        try
                        {
                            SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias57, v80 => { byrefalias57 = v80; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlContext = byrefalias57; }
                        _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter")), SWInstCounter), this, pDict, null, _.ARGS.Val("SumInstLicCounter"));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SET(_.SUBT(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter"))), this, pDict, null, _.ARGS.Val("SumFreeLicCounter"));

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
                    object byrefalias58 = hlContext;
                    try
                    {
                        SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias58, v81 => { byrefalias58 = v81; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlContext = byrefalias58; }
                    _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));

                    object byrefalias59 = hlContext;
                    try
                    {
                        SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias59, v82 => { byrefalias59 = v82; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlContext = byrefalias59; }
                    if (_.IF(_.GT(SWInstCounter, _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter")))))
                    {
                        _.SET(_.VAL(SWInstCounter), this, pDict, null, _.ARGS.Val("SumInstLicCounter"));
                    }
                }
                if (_.IF(_.EQ(_.NullableSTR(objType), "SoftwareLicense")))
                {
                    lstLicStatus = _.VAL(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseDetail.LicenseStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(lstLicStatus), "LicenseStatusValid")))
                    {
                        object byrefalias60 = hlContext;
                        try
                        {
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias60, v83 => { byrefalias60 = v83; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlContext = byrefalias60; }
                        _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SET(_.SUBT(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), _.CALLm0argp(this, pDict, _.ARGS.Val("SumInstLicCounter"))), this, pDict, null, _.ARGS.Val("SumFreeLicCounter"));
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
            cAssociationChanges = _.VAL(_.CALLm1v(this, hlContext, "GetAssociationChangesCount"));

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

                    AscDefNameChange = _.VAL(_.CALLm1v(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v(this, oAssociationChange, "IsToDelete")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v(this, hlObjB, "GetID"), _.CALLm2v(this, oAssociationChange, "EndB", "GetID"))))
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
            cAssociationChanges = _.VAL(_.CALLm1v(this, hlContext, "GetAssociationChangesCount"));

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

                    AscDefNameChange = _.VAL(_.CALLm1v(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v(this, oAssociationChange, "IsToCreate")))
                    {
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v(this, hlObjB, "GetID"), _.CALLm2v(this, oAssociationChange, "EndB", "GetID"))))
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
            cAssociationChanges = _.VAL(_.CALLm1v(this, hlContext, "GetAssociationChangesCount"));

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
                    AscDefNameChange = _.VAL(_.CALLm1v(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v(this, oAssociationChange, "IsToCreate")))
                    {
                        //ueberpruefen ob die gewuenschte Assoc auch angelegt werden soll.
                        if (_.IF(_.EQ(AscDefNameChange, AscDefName)))
                        {
                            if (_.IF(_.EQ(_.CALLm1v(this, hlObjB, "GetID"), _.CALLm2v(this, oAssociationChange, "EndB", "GetID"))))
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
            object byrefalias61 = pDict;
            try
            {
                rsltParent = _.VAL(_.CALLm1argp(this, hlChild, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).RefIfArray(byrefalias61, _.ARGS.Val("AssocID"))));
            }
            finally { pDict = byrefalias61; }
            if (_.IF(_.GTE(_.UBOUND(rsltParent), _.CALLm0argp(this, pDict, _.ARGS.Val("ParentCounter")))))
            {
                objParent = VBScriptConstants.Nothing;
                var enumerationContent5 = _.ENUMERABLE(rsltParent).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent5.MoveNext())
                        break;
                    objParent = enumerationContent5.Current;
                    object byrefalias62 = pDict;
                    try
                    {
                        GetAssociatedOrganizationalUnit_retVal = _.VAL(_.CALLm1argp(this, objParent, "GetValue", _.ARGS.Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias62, _.ARGS.Val("AttrName")).Val((Int16)0)));
                    }
                    finally { pDict = byrefalias62; }
                    object byrefalias63 = lcid;
                    try
                    {
                        outParentDefName = _.VAL(_.CALLm1argp(this, hlContext, "GetDisplayName", _.ARGS.Val(_.CALLm1argp(this, objParent, "GetValue", _.ARGS.Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("HLOBJECTINFO.DEFID"))).Ref(byrefalias63, v87 => { byrefalias63 = v87; })));
                    }
                    finally { lcid = byrefalias63; }
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
