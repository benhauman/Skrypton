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
                _.CALLm1v2(this, hlContext, "trace", (Int16)1, byrefalias6);
            }
            finally { text = byrefalias6; }
        }

        //---------------------------------------------------------------
        //Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
        public void SetPersonInformation(ref object hlContext, ref object hlObject, ref object dict)
        {
            object AttrDef = null;
            object strAttrValue = null;
            //Aus dem Dictionary wird das Attribut und der dazugehörige Wert ermittelt.
            AttrDef = "";
            AttrDef = _.CONCAT("PersonInformation.", _.CALLm0argp(this, dict, _.ARGS.Val("PersInfoAttr")));

            strAttrValue = "";
            strAttrValue = _.VAL(_.CALLm0argp(this, dict, _.ARGS.Val("DefValue")));

            if (_.IF(_.EQ(_.NullableSTR(strAttrValue), "")))
            {
                strAttrValue = "-";
            }
            _.CALLm1argp(this, hlObject, "SetValue", _.ARGS.Ref(AttrDef, v5 => { AttrDef = v5; }).Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strAttrValue, v6 => { strAttrValue = v6; }));
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
                _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(byrefalias7, v7 => { byrefalias7 = v7; }).Val(_.CONCAT("Type ", _.VARTYPE(hlObject))));
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
                strOrgUnits = _.VAL(_.CALLm1argp(this, hlOrgUnit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            else
            {
                strOrgUnits = _.CONCAT(strOrgUnits, ", ", _.CALLm1argp(this, hlOrgUnit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            orgaType = "";
            orgaType = _.VAL(_.CALLm1v0(this, hlOrgUnit, "GetType"));
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
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias8, v8 => { byrefalias8 = v8; }).RefIfArray(NextOrgUnit, _.ARGS.Val((Int16)0)).Ref(byrefalias9, v9 => { byrefalias9 = v9; })));
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
        //Über diese Function wird für ein Flag Attribut immer der Wert
        //True oder False ausgegeben.
        public object GetFlagValue(ref object hlContext, ref object hlObject, ref object hlattribute, ref object hlcontentid, ref object hlsuid)
        {
            object GetFlagValue_retVal = null;
            object byrefalias10 = hlattribute, byrefalias11 = hlcontentid, byrefalias12 = hlsuid;
            try
            {
                GetFlagValue_retVal = _.VAL(_.CALLm1argp(this, hlObject, "GetValue", _.ARGS.Ref(byrefalias10, v10 => { byrefalias10 = v10; }).Val((Int16)0).Ref(byrefalias11, v11 => { byrefalias11 = v11; }).Ref(byrefalias12, v12 => { byrefalias12 = v12; }).Val((Int16)0)));
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
        //Wörterbuch ohne Parameter.
        public object GetErrMsg0(ref object hlContext, ref object LocaleID, ref object ErrCode)
        {
            object GetErrMsg0_retVal = null;
            object strErrMsg = null;
            GetErrMsg0_retVal = "";

            strErrMsg = "";
            object byrefalias13 = ErrCode, byrefalias14 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias13, v13 => { byrefalias13 = v13; }).Ref(byrefalias14, v14 => { byrefalias14 = v14; })));
            }
            finally { ErrCode = byrefalias13; LocaleID = byrefalias14; }
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
            FirstOrgUnit = _.OBJ(_.CALLm1v0(this, hlContext, "GetRelatedObject"));

            bool ifResult;
            object byrefalias15 = hlContext;
            try
            {
                ifResult = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias15, v17 => { byrefalias15 = v17; }).Ref(FirstOrgUnit, v18 => { FirstOrgUnit = v18; })), true));
            }
            finally { hlContext = byrefalias15; }
            if (ifResult)
            {
                if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.CALLm1v0(this, FirstOrgUnit, "GetType")), "Company"), _.NOTEQ(_.NullableSTR(_.CALLm1v0(this, FirstOrgUnit, "GetType")), "Division"))))
                {
                    FirstOrgUnit = VBScriptConstants.Nothing;
                }
            }

            bool ifResult2;
            object byrefalias16 = hlContext;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias16, v21 => { byrefalias16 = v21; }).Ref(FirstOrgUnit, v22 => { FirstOrgUnit = v22; })), false));
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
                ifResult3 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias17, v25 => { byrefalias17 = v25; }).Ref(FirstOrgUnit, v26 => { FirstOrgUnit = v26; })), true));
            }
            finally { hlContext = byrefalias17; }
            if (ifResult3)
            {
                bool ifResult4;
                object byrefalias18 = hlContext;
                try
                {
                    ifResult4 = _.IF(_.EQ(_.NullableSTR(_.CALLm1argp(this, _outer, "GetBaseType", _.ARGS.Ref(byrefalias18, v29 => { byrefalias18 = v29; }).Ref(FirstOrgUnit, v30 => { FirstOrgUnit = v30; }))), "ORGANISATION"));
                }
                finally { hlContext = byrefalias18; }
                if (ifResult4)
                {
                    retval = "";
                    strOrgUnits = "";
                    object byrefalias19 = hlContext;
                    try
                    {
                        retval = _.VAL(_.CALLm1argp(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias19, v31 => { byrefalias19 = v31; }).Ref(FirstOrgUnit, v32 => { FirstOrgUnit = v32; }).Ref(strOrgUnits, v33 => { strOrgUnits = v33; })));
                    }
                    finally { hlContext = byrefalias19; }

                    _.SET(_.VAL(strOrgUnits), this, dict, null, _.ARGS.Val("DefValue"));
                    _.SET("PersonOrganization", this, dict, null, _.ARGS.Val("PersInfoAttr"));
                    object byrefalias20 = hlContext, byrefalias21 = hlPerson, byrefalias22 = dict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "SetPersonInformation", _.ARGS.Ref(byrefalias20, v34 => { byrefalias20 = v34; }).Ref(byrefalias21, v35 => { byrefalias21 = v35; }).Ref(byrefalias22, v36 => { byrefalias22 = v36; }));
                    }
                    finally { hlContext = byrefalias20; hlPerson = byrefalias21; dict = byrefalias22; }
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

            productDefName = _.VAL(_.CALLm1argp(this, hlObjectB, "GetType", _.ARGS.ForceBrackets()));

            if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(productDefName), "Software"), _.NOTEQ(_.NullableSTR(productDefName), "SoftwareLicence"))))
            {
                contentID = _.VAL(_.CALLm1argp(this, hlObjectB, "GenerateContentID", _.ARGS.ForceBrackets()));
                agentID = _.VAL(_.CALLm1argp(this, hlContext, "GetAgentID", _.ARGS.ForceBrackets()));
                orgUnitName = _.VAL(_.CALLm1argp(this, hlObjectA, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                personOfAgent = _.OBJ(_.CALLm1argp(this, hlContext, "GetPersonOfAgent", _.ARGS.Ref(agentID, v37 => { agentID = v37; })));
                if (_.IF(_.IS(personOfAgent, VBScriptConstants.Nothing)))
                {
                    object byrefalias23 = hlContext;
                    try
                    {
                        strErrMsg = _.VAL(_.CALLm1argp(this, _outer, "GetErrMsg0", _.ARGS.Ref(byrefalias23, v38 => { byrefalias23 = v38; }).Val(_.CALLm1v0(this, byrefalias23, "GetLocaleID")).Val("#ERR_SETASSETHISTORY")));
                    }
                    finally { hlContext = byrefalias23; }
                    object byrefalias24 = hlContext;
                    try
                    {
                        _.CALLm1argp(this, _outer, "Trace", _.ARGS.Ref(byrefalias24, v39 => { byrefalias24 = v39; }).Ref(strErrMsg, v40 => { strErrMsg = v40; }));
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
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedBy").Val((Int16)0).Ref(contentID, v41 => { contentID = v41; }).Val((Int16)0).Ref(personName, v42 => { personName = v42; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID").Val((Int16)0).Ref(contentID, v43 => { contentID = v43; }).Val((Int16)0).Ref(agentID, v44 => { agentID = v44; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangeDate").Val((Int16)0).Ref(contentID, v45 => { contentID = v45; }).Val((Int16)0).Val(_.NOW()));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnit").Val((Int16)0).Ref(contentID, v46 => { contentID = v46; }).Val((Int16)0).Ref(orgUnitName, v47 => { orgUnitName = v47; }));
                _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnitID").Val((Int16)0).Ref(contentID, v48 => { contentID = v48; }).Val((Int16)0).Val(_.CALLm1argp(this, hlObjectA, "GetID", _.ARGS.ForceBrackets())));

                if (_.IF(_.EQ(created, true)))
                {
                    _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v49 => { contentID = v49; }).Val((Int16)0).Val("HistoryActionCreated"));
                }
                else
                {
                    _.CALLm1argp(this, hlObjectB, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentID, v50 => { contentID = v50; }).Val((Int16)0).Val("HistoryActionDeleted"));
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
            object byrefalias25 = ErrCode, byrefalias26 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias25, v51 => { byrefalias25 = v51; }).Ref(byrefalias26, v52 => { byrefalias26 = v52; })));
            }
            finally { ErrCode = byrefalias25; LocaleID = byrefalias26; }
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
            object byrefalias27 = ErrCode, byrefalias28 = LocaleID;
            try
            {
                strErrMsg = _.VAL(_.CALLm1argp(this, hlContext, "GetTranslation", _.ARGS.Ref(byrefalias27, v53 => { byrefalias27 = v53; }).Ref(byrefalias28, v54 => { byrefalias28 = v54; })));
            }
            finally { ErrCode = byrefalias27; LocaleID = byrefalias28; }
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
            object byrefalias29 = HLASC_SoftwareLicenseFolderView;
            try
            {
                rsltSWFolders = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias29, v55 => { byrefalias29 = v55; })));
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
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias30, v56 => { byrefalias30 = v56; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias31, v57 => { byrefalias31 = v57; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias32, v58 => { byrefalias32 = v58; }).Ref(byrefalias33, v59 => { byrefalias33 = v59; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
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
                NextSWFolder = _.VAL(_.CALLm1argp(this, hlParentSWFolder, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias34, v60 => { byrefalias34 = v60; })));
            }
            finally { HLASC_SoftwareLicenseFolderView = byrefalias34; }
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(NextSWFolder)), (Int16)0)))
            {
                object byrefalias35 = hlContext, byrefalias36 = pDict, byrefalias37 = HLASC_SoftwareLicenseFolderView;
                try
                {
                    retval = _.VAL(_.CALLm1argp(this, _outer, "CheckForSoftwareSuiteFolder", _.ARGS.Ref(byrefalias35, v61 => { byrefalias35 = v61; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(byrefalias36, v62 => { byrefalias36 = v62; }).Ref(byrefalias37, v63 => { byrefalias37 = v63; })));
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

            //Dictionary Einträge initalisieren
            _.SET("", this, pDict, null, _.ARGS.Val("SoftwareLicenses"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumInstLicCounter"));
            _.SET((Int16)0, this, pDict, null, _.ARGS.Val("SumFreeLicCounter"));

            //Prüfen ob es Software Lizenzobjekte unterhalb des Folders gibt.
            object byrefalias38 = assocName;
            try
            {
                _.SET(_.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias38, v64 => { byrefalias38 = v64; }))), this, pDict, null, _.ARGS.Val("SoftwareLicenses"));
            }
            finally { assocName = byrefalias38; }

            //Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
            //Objekte gezählt werden müssen
            CheckSoftwareSuite = false;
            object byrefalias39 = hlContext, byrefalias40 = hlSWFolder;
            try
            {
                CheckSoftwareSuite = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias39, v65 => { byrefalias39 = v65; }).Ref(byrefalias40, v66 => { byrefalias40 = v66; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
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
                        _.CALLm1argp(this, _outer, "CalcAllLicCounter", _.ARGS.Ref(byrefalias42, v67 => { byrefalias42 = v67; }).Ref(byrefalias43, v68 => { byrefalias43 = v68; }));
                    }
                    finally { hlContext = byrefalias42; pDict = byrefalias43; }
                }
                else
                {
                    object byrefalias44 = hlContext, byrefalias45 = pDict;
                    try
                    {
                        _.CALLm1argp(this, _outer, "CalcFolderLicCounter", _.ARGS.Ref(byrefalias44, v69 => { byrefalias44 = v69; }).Ref(byrefalias45, v70 => { byrefalias45 = v70; }));
                    }
                    finally { hlContext = byrefalias44; pDict = byrefalias45; }
                }
            }
            //Gesatmzahl der Lizenzen in den Lizenzumschlag zurückschreiben
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
                CheckLicContrByServer = _.VAL(_.CALLm1argp(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias48, v71 => { byrefalias48 = v71; }).Ref(byrefalias49, v72 => { byrefalias49 = v72; }).Val("SoftwareLicenseFolderDetail.FlagLicenseControlledByServer").Val((Int16)0).Val((Int16)0)));
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

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            NextSWFolder = "";
            a = "";
            a = _.VAL(_.CALLm1v0(this, hlSWFolder, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(a), "LicenseFolder")))
            {
                object byrefalias51 = assocName;
                try
                {
                    NextSWFolder = _.VAL(_.CALLm1argp(this, hlSWFolder, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Ref(byrefalias51, v73 => { byrefalias51 = v73; })));
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
                    retval = _.VAL(_.CALLm1argp(this, _outer, "SetLicenseCounter", _.ARGS.Ref(byrefalias52, v74 => { byrefalias52 = v74; }).RefIfArray(NextSWFolder, _.ARGS.Val((Int16)0)).Ref(byrefalias53, v75 => { byrefalias53 = v75; }).Ref(byrefalias54, v76 => { byrefalias54 = v76; })));
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
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias55, v77 => { byrefalias55 = v77; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias56, v78 => { byrefalias56 = v78; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlContext = byrefalias56; }
                        _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));
                        object byrefalias57 = hlContext;
                        try
                        {
                            SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias57, v79 => { byrefalias57 = v79; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                        SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias58, v80 => { byrefalias58 = v80; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlContext = byrefalias58; }
                    _.SET(_.ADD(_.CALLm0argp(this, pDict, _.ARGS.Val("SumRefLicCounter")), SWRefLicCounter), this, pDict, null, _.ARGS.Val("SumRefLicCounter"));

                    object byrefalias59 = hlContext;
                    try
                    {
                        SWInstCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias59, v81 => { byrefalias59 = v81; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                            SWRefLicCounter = _.VAL(_.CALLm1argp(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias60, v82 => { byrefalias60 = v82; }).Val(_.CALLm1argp(this, SoftwareLicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
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
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v83 => { ixAC = v83; })));

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
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v84 => { ixAC = v84; })));

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

            //Anzahl der zu erstellenden oder löschenden Assoziationen
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

                    //Für jede Assoziations Änderung wird das entsprechende Infos (Objekt    ) ausgelsen.
                    oAssociationChange = _.OBJ(_.CALLm1argp(this, hlContext, "GetAssociationChangeAt", _.ARGS.Ref(ixAC, v85 => { ixAC = v85; })));
                    //Def Name der Assoc ermitteln, die angelegt werden soll
                    AscDefNameChange = _.VAL(_.CALLm1v0(this, oAssociationChange, "AssociationType"));

                    if (_.IF(_.CALLm1v0(this, oAssociationChange, "IsToCreate")))
                    {
                        //Überprüfen ob die gewünschte Assoc auch angelegt werden soll.
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
                        GetAssociatedOrganizationalUnit_retVal = _.VAL(_.CALLm1argp(this, objParent, "GetValue", _.ARGS.RefIfArray(byrefalias62, _.ARGS.Val("AttrName")).Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { pDict = byrefalias62; }
                    object byrefalias63 = lcid;
                    try
                    {
                        outParentDefName = _.VAL(_.CALLm1v2(this, hlContext, "GetDisplayName", _.CALLm1argp(this, objParent, "GetValue", _.ARGS.Val("HLOBJECTINFO.DEFID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)), byrefalias63));
                    }
                    finally { lcid = byrefalias63; }
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
            objXMLDoc = _.OBJ(_.CREATEOBJECT("Msxml2.DOMDocument"));

            //XML-Processing Instruction hinzufügen
            xmlProInc = VBScriptConstants.Nothing;
            xmlProInc = _.OBJ(_.CALLm1v2(this, objXMLDoc, "createProcessingInstruction", "xml", "version='1.0' encoding='UTF-8'"));
            _.CALLm1argp(this, objXMLDoc, "insertBefore", _.ARGS.Ref(xmlProInc, v86 => { xmlProInc = v86; }).Val(_.CALLm1v0(this, objXMLDoc, "firstChild")));

            //Root-Element erstellen
            xmlRoot = _.OBJ(_.CALLm1v1(this, objXMLDoc, "CreateElement", "ASAPBatch"));
            _.CALLm1v1(this, objXMLDoc, "AppendChild", xmlRoot);
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "xmlns", "http://www.brainware.ch/operationsmanager/asap-batch/1.1");
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "xmlns:dt", "http://www.brainware.ch/operationsmanager/wf/changemanagement/columbus/datatypes/1.1");
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "xsi:schemaLocation", "http://www.brainware.ch/operationsmanager/asap-batch/1.1 asap-batch-1.1.xsd");
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "version", "1.1");
            _.CALLm1v2(this, xmlRoot, "SetAttribute", "responseRequired", "Yes");

            //Das Node Session hinzufügen
            nodeSession = _.OBJ(_.CALLm1v1(this, objXMLDoc, "CreateElement", "Session"));
            _.CALLm1v1(this, xmlRoot, "AppendChild", nodeSession);
            _.CALLm1v2(this, nodeSession, "SetAttribute", "id", "s1");
            _.CALLm1v2(this, nodeSession, "SetAttribute", "loginname", "foreignSystems\\assetcolumbus");
            _.CALLm1v2(this, nodeSession, "SetAttribute", "password", "");

            //XML Dokument inkl. Header an das Dictionary übergeben
            _.SET(_.OBJ(objXMLDoc), this, pDict, null, _.ARGS.Val("XMLDocument"));
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
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, xmlRoot, "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "id", "e7");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_adddevice");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeObserverKey);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("ObserverKey"))), this, nodeObserverKey, "Text");

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ContextData"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeAddDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:AddDeviceActualParams"));
            _.CALLm1v1(this, nodeContextData, "AppendChild", nodeAddDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDeviceName);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("DeviceName"))), this, nodeDeviceName, "Text");

            //Das Node CompanyName hinzufügen
            nodeCmpyName = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:CompanyName"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeCmpyName);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("CompanyName"))), this, nodeCmpyName, "Text");

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDomain);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("Domain"))), this, nodeDomain, "Text");

            //Das Node CostCenter hinzufügen
            nodeCostCenter = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:CostCenter"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeCostCenter);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("CostCenter"))), this, nodeCostCenter, "Text");

            //Das Node MACAdess hinzufügen
            nodeMACAddress = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:MACAddress"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeMACAddress);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("MACAddress"))), this, nodeMACAddress, "Text");

            //Das Node SubnetMask hinzufügen
            nodeSubnetMask = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:SubnetMask"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeSubnetMask);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("SubnetMask"))), this, nodeSubnetMask, "Text");

            //Das Node HwTypeId hinzufügen
            nodeHWType = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:HwTypeId"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeHWType);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("HwTypeId"))), this, nodeHWType, "Text");

            //Das Node OsTypeId hinzufügen
            nodeOSType = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:OsTypeId"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeOSType);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("OsTypeId"))), this, nodeOSType, "Text");

            //Das Node ActivationState hinzufügen
            nodeActState = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:ActivationState"));
            _.CALLm1v1(this, nodeAddDeviceActualParams, "AppendChild", nodeActState);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("ActivationState"))), this, nodeActState, "Text");

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
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, xmlRoot, "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "id", "e7");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_chgdevice");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeObserverKey);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("ObserverKey"))), this, nodeObserverKey, "Text");

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ContextData"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeChgDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:ChangeDeviceActualParams"));
            _.CALLm1v1(this, nodeContextData, "AppendChild", nodeChgDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDeviceName);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("DeviceName"))), this, nodeDeviceName, "Text");

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDomain);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("Domain"))), this, nodeDomain, "Text");

            //Das Node CompanyName hinzufügen
            nodeCmpyName = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:CompanyName"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeCmpyName);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("CompanyName"))), this, nodeCmpyName, "Text");

            //Das Node CostCenter hinzufügen
            nodeCostCenter = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:CostCenter"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeCostCenter);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("CostCenter"))), this, nodeCostCenter, "Text");

            //Das Node MACAdess hinzufügen
            nodeMACAddress = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:MACAddress"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeMACAddress);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("MACAddress"))), this, nodeMACAddress, "Text");

            //Das Node SubnetMask hinzufügen
            nodeSubnetMask = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:SubnetMask"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeSubnetMask);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("SubnetMask"))), this, nodeSubnetMask, "Text");

            //Das Node HwTypeId hinzufügen
            nodeHWType = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:HwTypeId"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeHWType);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("HwTypeId"))), this, nodeHWType, "Text");

            //Das Node OsTypeId hinzufügen
            nodeOSType = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:OsTypeId"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeOSType);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("OsTypeId"))), this, nodeOSType, "Text");

            //Das Node ActivationState hinzufügen
            nodeActState = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:ActivationState"));
            _.CALLm1v1(this, nodeChgDeviceActualParams, "AppendChild", nodeActState);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("ActivationState"))), this, nodeActState, "Text");

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
            xmlRoot = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodeCreateInstanceRq = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "CreateInstanceRq"));
            _.CALLm1v1(this, xmlRoot, "AppendChild", nodeCreateInstanceRq);
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "id", "e7");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfpNs", "ch.bw.wf.changemgmt.columbus_removedevice");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "wfmNs", "Columbus Changemanagement");
            _.CALLm1v2(this, nodeCreateInstanceRq, "SetAttribute", "sessionId", "s1");

            //Das Node ObserverKey hinzufügen
            nodeObserverKey = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ObserverKey"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeObserverKey);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("ObserverKey"))), this, nodeObserverKey, "Text");

            //Das Container Node ContextData hinzufügen
            nodeContextData = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "ContextData"));
            _.CALLm1v1(this, nodeCreateInstanceRq, "AppendChild", nodeContextData);

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeRemoveDeviceActualParams = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:RemoveDeviceActualParams"));
            _.CALLm1v1(this, nodeContextData, "AppendChild", nodeRemoveDeviceActualParams);

            //Das Container Node DeviceIdentification hinzufügen
            nodeDeviceIdentification = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceIdentification"));
            _.CALLm1v1(this, nodeRemoveDeviceActualParams, "AppendChild", nodeDeviceIdentification);

            //Das Node DeviceName hinzufügen
            nodeDeviceName = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:DeviceName"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDeviceName);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("DeviceName"))), this, nodeDeviceName, "Text");

            //Das Node CompanyName hinzufügen
            //Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
            //nodeDeviceIdentification.AppendChild (nodeCmpyName)
            //nodeCmpyName.Text = pDict("CompanyName")

            //Das Node Domain hinzufügen
            nodeDomain = _.OBJ(_.CALLm1v1(this, _.CALLm0argp(this, pDict, _.ARGS.Val("XMLDocument")), "CreateElement", "dt:Domain"));
            _.CALLm1v1(this, nodeDeviceIdentification, "AppendChild", nodeDomain);
            _.SET(_.VAL(_.CALLm0argp(this, pDict, _.ARGS.Val("Domain"))), this, nodeDomain, "Text");

            return MIG_CreateDELXML2Columbus_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //Wenn beide Werte ein Datum sind, muss geprüft werden ob das Enddatum nach dem
        //Start Datum liegt. Falls nicht wird "False" zurückgegeben.
        public object MigCheckDatePeriod(ref object hlContext, ref object StartDate, ref object EndDate)
        {
            object MigCheckDatePeriod_retVal = null;
            MigCheckDatePeriod_retVal = false;

            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALLm1v2(this, _, "DATEPART", "d", _.CDATE(StartDate))), "0")))
            {
                if (_.IF(_.LT(_.CALLm1v2(this, _, "DATEPART", "d", _.CDATE(StartDate)), _.CALLm1v2(this, _, "DATEPART", "d", _.CDATE(EndDate)))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    MigCheckDatePeriod_retVal = true;
                }

                if (_.IF(_.GT(_.CALLm1v2(this, _, "DATEPART", "yyyy", _.CDATE(StartDate)), _.CALLm1v2(this, _, "DATEPART", "yyyy", _.CDATE(EndDate)))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    if (_.IF(_.GT(_.CALLm1v2(this, _, "DATEPART", "y", _.CDATE(StartDate)), _.CALLm1v2(this, _, "DATEPART", "y", _.CDATE(EndDate)))))
                    {
                        if (_.IF(_.LT(_.CALLm1v2(this, _, "DATEPART", "yyyy", _.CDATE(StartDate)), _.CALLm1v2(this, _, "DATEPART", "yyyy", _.CDATE(EndDate)))))
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
            Qry = _.OBJ(_.CALLm1argp(this, hlSrvContext, "OpenSearch", _.ARGS.Ref(srchQuery, v87 => { srchQuery = v87; })));
            rsltQuery = _.VAL(_.CALLm1argp(this, Qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
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
            intAgentID = _.VAL(_.CALLm1argp(this, hlContext, "GetAgentID", _.ARGS.ForceBrackets()));
            objPerson = VBScriptConstants.Nothing;

            objPerson = _.OBJ(_.CALLm1argp(this, hlContext, "GetPersonOfAgent", _.ARGS.Ref(intAgentID, v88 => { intAgentID = v88; })));

            bool ifResult6;
            object byrefalias64 = hlContext;
            try
            {
                ifResult6 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias64, v91 => { byrefalias64 = v91; }).Ref(objPerson, v92 => { objPerson = v92; })), true));
            }
            finally { hlContext = byrefalias64; }
            if (ifResult6)
            {

                if (_.IF(_.NOTEQ(_.NullableSTR(relObjMIGPartnerID), "")))
                {

                    strPersonInternalMIGPartnerIDs = _.VAL(_.CALLm1argp(this, objPerson, "GetValue", _.ARGS.Val("MIGAgentInformation.InternalMIGPartnerID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

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
