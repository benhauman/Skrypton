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

            _outer.hlasc_software2computer = "Software2Computer";
            _outer.hlasc_softwarelicensegroupview = "LicenseGroupView";
            _outer.hlasc_softwarelicensefolderview = "LicenseFolderView";
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
            hlasc_softwarelicensefolderview = null;
            hlasc_softwarelicensegroupview = null;
            hlasc_software2computer = null;
        }

        internal object hlasc_softwarelicensefolderview { get; set; }
        internal object hlasc_softwarelicensegroupview { get; set; }
        internal object hlasc_software2computer { get; set; }

        //---------------------------------------------------------------
        //Diese Funktion ermittelt den Standard-Eintrag zum angegebenen Attribut aus
        //dem Dictionary.
        //Wenn der Parameter "GetAll" auf False steht wird als Rückgabewert für die Funktion
        //ebenfalls "False" ausgegben, wenn mehr als ein Standardeintrag gefunden wird.
        //Wenn für den Parameter "True" angeben wird, prüft die Funktion ob es tatsächlich
        //nur einen Standard-Eintrag gibt, sonst "False".
        public object getcommunicationdefault(ref object hlcontext, ref object hlobject, ref object dict, ref object getall)
        {
            object GetCommunicationDefault_retVal = null;
            object itemcount = null;
            object strvalue = null;
            object itemids = null;
            object item = null;
            object defitem = null;
            GetCommunicationDefault_retVal = false;
            itemcount = (Int16)0;
            strvalue = "";

            itemids = "";
            object byrefalias = dict;
            try
            {
                itemids = _.VAL(_.CALL(this, hlobject, "GetContentIDs", _.ARGS.RefIfArray(byrefalias, _.ARGS.Val("Compound")).Val((Int16)0)));
            }
            finally { dict = byrefalias; }

            item = (Int16)0;
            var enumerationContent = _.ENUMERABLE(itemids).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                defitem = false;
                object byrefalias2 = hlcontext, byrefalias3 = hlobject, byrefalias4 = dict;
                try
                {
                    defitem = _.VAL(_.CALL(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias2, v => { byrefalias2 = v; }).Ref(byrefalias3, v2 => { byrefalias3 = v2; }).RefIfArray(byrefalias4, _.ARGS.Val("Default")).Ref(item, v3 => { item = v3; }).Val((Int16)0)));
                }
                finally { hlcontext = byrefalias2; hlobject = byrefalias3; dict = byrefalias4; }
                if (_.IF(_.EQ(_.CBOOL(defitem), true)))
                {
                    itemcount = _.ADD(itemcount, (Int16)1);
                    object byrefalias5 = dict;
                    try
                    {
                        strvalue = _.VAL(_.CALL(this, hlobject, "GetValue", _.ARGS.RefIfArray(byrefalias5, _.ARGS.Val("Value")).Val((Int16)0).Ref(item, v4 => { item = v4; }).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { dict = byrefalias5; }
                    if (_.IF(_.EQ(_.CBOOL(getall), false)))
                    {
                        break;
                    }
                }
            }
            if (_.IF(_.GT(_.NullableNUM(itemcount), (Int16)1)))
            {
                GetCommunicationDefault_retVal = false;
                return GetCommunicationDefault_retVal;
            }
            else
            {
                GetCommunicationDefault_retVal = true;
                _.SET(_.VAL(strvalue), this, dict, null, _.ARGS.Val("DefValue"));
            }
            return GetCommunicationDefault_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
        public void trace(ref object hlcontext, ref object text)
        {
            object byrefalias6 = text;
            try
            {
                _.CALL(this, hlcontext, "trace", _.ARGS.Val((Int16)1).Ref(byrefalias6, v5 => { byrefalias6 = v5; }));
            }
            finally { text = byrefalias6; }
        }

        //---------------------------------------------------------------
        //Setzt den vorhandenen Wert aus dem VB-Dictionary in die ODE "PersonInformation".
        public void setpersoninformation(ref object hlcontext, ref object hlobject, ref object dict)
        {
            object attrdef = null;
            object strattrvalue = null;
            //Aus dem Dictionary wird das Attribut und der dazugehörige Wert ermittelt.
            attrdef = "";
            attrdef = _.CONCAT("PersonInformation.", _.CALL(this, dict, _.ARGS.Val("PersInfoAttr")));

            strattrvalue = "";
            strattrvalue = _.VAL(_.CALL(this, dict, _.ARGS.Val("DefValue")));

            if (_.IF(_.EQ(_.NullableSTR(strattrvalue), "")))
            {
                strattrvalue = "-";
            }
            _.CALL(this, hlobject, "SetValue", _.ARGS.Ref(attrdef, v6 => { attrdef = v6; }).Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(strattrvalue, v7 => { strattrvalue = v7; }));
        }

        //---------------------------------------------------------------
        public object ishlobject(ref object hlcontext, ref object hlobject)
        {
            object IsHLObject_retVal = null;
            //	Trace hlContext, "IsObject " & IsObject(hlObject)
            //	Trace hlContext, "IsNull " & IsNull(hlObject)
            //	Trace hlContext, "IsEmpty " & IsEmpty(hlObject)
            //	Trace hlContext, "Leerstring "
            //	Trace hlContext, "Leerstring " & hlObject = ""
            object byrefalias7 = hlcontext;
            try
            {
                _.CALL(this, _outer, "Trace", _.ARGS.Ref(byrefalias7, v8 => { byrefalias7 = v8; }).Val(_.CONCAT("Type ", _.VARTYPE(hlobject))));
            }
            finally { hlcontext = byrefalias7; }
            IsHLObject_retVal = _.VAL(_.AND(_.EQ(_.ISOBJECT(hlobject), true), _.EQ(_.IS(hlobject, VBScriptConstants.Nothing), false)));
            return IsHLObject_retVal;
        }

        //-------------------------------------------------------------------
        public object getbasetype(ref object hlcontext, ref object hlobject)
        {
            return _.VAL(_.CALL(this, hlobject, "GetValue", _.ARGS.Val("HLOBJECTINFO.BASETYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
        }

        //---------------------------------------------------------------
        //Dies ist eine rekursive Function zum ermitteln der Organisationshierarchie,
        //ausgehend vom der ersten OU überhalb einer Person.
        //Die Variable "strOrgUnits" ist der Out-Parameter der Function.
        public object getpersonorganisation(ref object hlcontext, ref object hlorgunit, ref object strorgunits)
        {
            object GetPersonOrganisation_retVal = null;
            object retval = null;
            object nextorgunit = null;
            object orgatype = null;
            GetPersonOrganisation_retVal = (Int16)0;
            retval = (Int16)0;

            //Wenn noch keine OU ermittelt wurde, wird der Name der ersten OU eingetragen.
            //Andernfalls, wird jede weitere OU einfach angehangen.
            if (_.IF(_.EQ(_.NullableSTR(strorgunits), "")))
            {
                strorgunits = _.VAL(_.CALL(this, hlorgunit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }
            else
            {
                strorgunits = _.CONCAT(strorgunits, ", ", _.CALL(this, hlorgunit, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            }

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            orgatype = "";
            orgatype = _.VAL(_.CALL(this, hlorgunit, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(orgatype), "Division")))
            {
                nextorgunit = _.VAL(_.CALL(this, hlorgunit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("CompanyView")));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgatype), "Site")))
            {
                nextorgunit = _.VAL(_.CALL(this, hlorgunit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Site2Company")));
            }
            if (_.IF(_.EQ(_.NullableSTR(orgatype), "Company")))
            {
                nextorgunit = _.VAL(_.CALL(this, hlorgunit, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Company2Company")));
            }

            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.ISARRAY(nextorgunit)))
            {
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(nextorgunit)), (Int16)0)))
                {
                    object byrefalias8 = hlcontext, byrefalias9 = strorgunits;
                    try
                    {
                        retval = _.VAL(_.CALL(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias8, v9 => { byrefalias8 = v9; }).RefIfArray(nextorgunit, _.ARGS.Val((Int16)0)).Ref(byrefalias9, v10 => { byrefalias9 = v10; })));
                    }
                    finally { hlcontext = byrefalias8; strorgunits = byrefalias9; }
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
        public object getflagvalue(ref object hlcontext, ref object hlobject, ref object hlattribute, ref object hlcontentid, ref object hlsuid)
        {
            object GetFlagValue_retVal = null;
            object byrefalias10 = hlattribute, byrefalias11 = hlcontentid, byrefalias12 = hlsuid;
            try
            {
                GetFlagValue_retVal = _.VAL(_.CALL(this, hlobject, "GetValue", _.ARGS.Ref(byrefalias10, v11 => { byrefalias10 = v11; }).Val((Int16)0).Ref(byrefalias11, v12 => { byrefalias11 = v12; }).Ref(byrefalias12, v13 => { byrefalias12 = v13; }).Val((Int16)0)));
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
        public object geterrmsg0(ref object hlcontext, ref object localeid, ref object errcode)
        {
            object GetErrMsg0_retVal = null;
            object strerrmsg = null;
            GetErrMsg0_retVal = "";

            strerrmsg = "";
            object byrefalias13 = errcode, byrefalias14 = localeid;
            try
            {
                strerrmsg = _.VAL(_.CALL(this, hlcontext, "GetTranslation", _.ARGS.Ref(byrefalias13, v14 => { byrefalias13 = v14; }).Ref(byrefalias14, v15 => { byrefalias14 = v15; })));
            }
            finally { errcode = byrefalias13; localeid = byrefalias14; }
            strerrmsg = _.CONCAT(strerrmsg, VBScriptConstants.vbNewLine, "(Code: ", errcode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg0_retVal = _.REPLACE(strerrmsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg0_retVal;
        }

        //Das Script ermittelt auf Basis der ersten übergeordneten OU den gesamten Pfad bis zur Firma oder Konzern
        //und speichert diesen in das Hilfsattribut PersonInformation.PersonOrganisation.
        //This script detects the entire path based on the first parent OU up to the company or holding
        //and saves them into the attribute PersonInformation.PersonOrganisation.
        public void setpersonorganization(ref object hlcontext, ref object hlperson, ref object dict)
        {
            object firstorgunit = null;
            object rsltorgunit = null;
            object retval = null;
            object strorgunits = null;
            firstorgunit = VBScriptConstants.Nothing;
            firstorgunit = _.OBJ(_.CALL(this, hlcontext, "GetRelatedObject"));

            bool ifResult;
            object byrefalias15 = hlcontext;
            try
            {
                ifResult = _.IF(_.EQ(_.CALL(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias15, v18 => { byrefalias15 = v18; }).Ref(firstorgunit, v19 => { firstorgunit = v19; })), true));
            }
            finally { hlcontext = byrefalias15; }
            if (ifResult)
            {
                if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.CALL(this, firstorgunit, "GetType")), "Company"), _.NOTEQ(_.NullableSTR(_.CALL(this, firstorgunit, "GetType")), "Division"))))
                {
                    firstorgunit = VBScriptConstants.Nothing;
                }
            }

            bool ifResult2;
            object byrefalias16 = hlcontext;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALL(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias16, v22 => { byrefalias16 = v22; }).Ref(firstorgunit, v23 => { firstorgunit = v23; })), false));
            }
            finally { hlcontext = byrefalias16; }
            if (ifResult2)
            {
                rsltorgunit = "";
                rsltorgunit = _.VAL(_.CALL(this, hlperson, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Val("Person2Organization")));
                if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltorgunit)), (Int16)0)))
                {
                    firstorgunit = _.OBJ(_.CALL(this, rsltorgunit, _.ARGS.Val((Int16)0)));
                }
            }

            bool ifResult3;
            object byrefalias17 = hlcontext;
            try
            {
                ifResult3 = _.IF(_.EQ(_.CALL(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias17, v26 => { byrefalias17 = v26; }).Ref(firstorgunit, v27 => { firstorgunit = v27; })), true));
            }
            finally { hlcontext = byrefalias17; }
            if (ifResult3)
            {
                bool ifResult4;
                object byrefalias18 = hlcontext;
                try
                {
                    ifResult4 = _.IF(_.EQ(_.NullableSTR(_.CALL(this, _outer, "GetBaseType", _.ARGS.Ref(byrefalias18, v30 => { byrefalias18 = v30; }).Ref(firstorgunit, v31 => { firstorgunit = v31; }))), "ORGANISATION"));
                }
                finally { hlcontext = byrefalias18; }
                if (ifResult4)
                {
                    retval = "";
                    strorgunits = "";
                    object byrefalias19 = hlcontext;
                    try
                    {
                        retval = _.VAL(_.CALL(this, _outer, "GetPersonOrganisation", _.ARGS.Ref(byrefalias19, v32 => { byrefalias19 = v32; }).Ref(firstorgunit, v33 => { firstorgunit = v33; }).Ref(strorgunits, v34 => { strorgunits = v34; })));
                    }
                    finally { hlcontext = byrefalias19; }

                    _.SET(_.VAL(strorgunits), this, dict, null, _.ARGS.Val("DefValue"));
                    _.SET("PersonOrganization", this, dict, null, _.ARGS.Val("PersInfoAttr"));
                    object byrefalias20 = hlcontext, byrefalias21 = hlperson, byrefalias22 = dict;
                    try
                    {
                        _.CALL(this, _outer, "SetPersonInformation", _.ARGS.Ref(byrefalias20, v35 => { byrefalias20 = v35; }).Ref(byrefalias21, v36 => { byrefalias21 = v36; }).Ref(byrefalias22, v37 => { byrefalias22 = v37; }));
                    }
                    finally { hlcontext = byrefalias20; hlperson = byrefalias21; dict = byrefalias22; }
                }
            }
        }

        //----------------------------------------------------------------------------------------------------------
        //Prozedur füllt die Umzugshistorie für das entsprechende Objekt
        public void setassethistory(ref object hlcontext, ref object hlobjecta, ref object hlobjectb, ref object created)
        {
            object productdefname = null;
            object agentid = null;
            object contentid = null;
            object personofagent = null;
            object personname = null;
            object orgunitname = null;
            object strerrmsg = null;

            productdefname = _.VAL(_.CALL(this, hlobjectb, "GetType", _.ARGS.ForceBrackets()));

            if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(productdefname), "Software"), _.NOTEQ(_.NullableSTR(productdefname), "SoftwareLicence"))))
            {
                contentid = _.VAL(_.CALL(this, hlobjectb, "GenerateContentID", _.ARGS.ForceBrackets()));
                agentid = _.VAL(_.CALL(this, hlcontext, "GetAgentID", _.ARGS.ForceBrackets()));
                orgunitname = _.VAL(_.CALL(this, hlobjecta, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                personofagent = _.OBJ(_.CALL(this, hlcontext, "GetPersonOfAgent", _.ARGS.Ref(agentid, v38 => { agentid = v38; })));
                if (_.IF(_.IS(personofagent, VBScriptConstants.Nothing)))
                {
                    object byrefalias23 = hlcontext;
                    try
                    {
                        strerrmsg = _.VAL(_.CALL(this, _outer, "GetErrMsg0", _.ARGS.Ref(byrefalias23, v39 => { byrefalias23 = v39; }).Val(_.CALL(this, byrefalias23, "GetLocaleID")).Val("#ERR_SETASSETHISTORY")));
                    }
                    finally { hlcontext = byrefalias23; }
                    object byrefalias24 = hlcontext;
                    try
                    {
                        _.CALL(this, _outer, "Trace", _.ARGS.Ref(byrefalias24, v40 => { byrefalias24 = v40; }).Ref(strerrmsg, v41 => { strerrmsg = v41; }));
                    }
                    finally { hlcontext = byrefalias24; }
                    //hlContext.abortcommand strErrMsg
                }
                else
                {
                    personname = _.VAL(_.CALL(this, personofagent, "GetValue", _.ARGS.Val("PersonGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    personname = _.CONCAT(personname, ", ");
                    personname = _.CONCAT(personname, _.CALL(this, personofagent, "GetValue", _.ARGS.Val("PersonGeneral.GivenName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                }
                _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedBy").Val((Int16)0).Ref(contentid, v42 => { contentid = v42; }).Val((Int16)0).Ref(personname, v43 => { personname = v43; }));
                _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangedByAgentID").Val((Int16)0).Ref(contentid, v44 => { contentid = v44; }).Val((Int16)0).Ref(agentid, v45 => { agentid = v45; }));
                _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryChangeDate").Val((Int16)0).Ref(contentid, v46 => { contentid = v46; }).Val((Int16)0).Val(_.NOW()));
                _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnit").Val((Int16)0).Ref(contentid, v47 => { contentid = v47; }).Val((Int16)0).Ref(orgunitname, v48 => { orgunitname = v48; }));
                _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryOrgUnitID").Val((Int16)0).Ref(contentid, v49 => { contentid = v49; }).Val((Int16)0).Val(_.CALL(this, hlobjecta, "GetID", _.ARGS.ForceBrackets())));

                if (_.IF(_.EQ(created, true)))
                {
                    _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentid, v50 => { contentid = v50; }).Val((Int16)0).Val("HistoryActionCreated"));
                }
                else
                {
                    _.CALL(this, hlobjectb, "SetValue", _.ARGS.Val("AssocHistory.HistoryInformation_CA.HistoryAction").Val((Int16)0).Ref(contentid, v51 => { contentid = v51; }).Val((Int16)0).Val("HistoryActionDeleted"));
                }
            }
        }

        //---------------------------------------------------------------
        //Diese Function ermitellt eine Fehlermeldung aus dem helpLine
        //Wörterbuch mit einem Parameter.
        public object geterrmsg1(ref object hlcontext, ref object localeid, ref object errcode, ref object arg1)
        {
            object GetErrMsg1_retVal = null;
            object strerrmsg = null;
            GetErrMsg1_retVal = "";

            strerrmsg = "";
            object byrefalias25 = errcode, byrefalias26 = localeid;
            try
            {
                strerrmsg = _.VAL(_.CALL(this, hlcontext, "GetTranslation", _.ARGS.Ref(byrefalias25, v52 => { byrefalias25 = v52; }).Ref(byrefalias26, v53 => { byrefalias26 = v53; })));
            }
            finally { errcode = byrefalias25; localeid = byrefalias26; }
            strerrmsg = _.REPLACE(strerrmsg, "%1", arg1);
            strerrmsg = _.CONCAT(strerrmsg, VBScriptConstants.vbLf, "(Code: ", errcode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg1_retVal = _.REPLACE(strerrmsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg1_retVal;
        }

        public object geterrmsg2(ref object hlcontext, ref object localeid, ref object errcode, ref object arg1, ref object arg2)
        {
            object GetErrMsg2_retVal = null;
            object strerrmsg = null;
            GetErrMsg2_retVal = "";

            strerrmsg = "";
            object byrefalias27 = errcode, byrefalias28 = localeid;
            try
            {
                strerrmsg = _.VAL(_.CALL(this, hlcontext, "GetTranslation", _.ARGS.Ref(byrefalias27, v54 => { byrefalias27 = v54; }).Ref(byrefalias28, v55 => { byrefalias28 = v55; })));
            }
            finally { errcode = byrefalias27; localeid = byrefalias28; }
            strerrmsg = _.REPLACE(strerrmsg, "%1", arg1);
            strerrmsg = _.REPLACE(strerrmsg, "%2", arg2);
            strerrmsg = _.CONCAT(strerrmsg, VBScriptConstants.vbLf, "(Code: ", errcode, ")");

            //Den Paramenter %LF% durch Zeilenumbrüche ersetzen.
            //Rückgabewert der Function ist die Fehlermeldung.
            GetErrMsg2_retVal = _.REPLACE(strerrmsg, "%LF%", VBScriptConstants.vbNewLine);
            return GetErrMsg2_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //In dieser Funktion wird geprüft, ob es unterhalb einer Software Suite
        //bereits Lizenzumschläge mit Lizenzen gibt.
        public object getreferencelicensecount(ref object hlcontext, ref object hlswfolder, ref object chkfolderonly, ref object hlasc_softwarelicensefolderview)
        {
            object GetReferenceLicenseCount_retVal = null;
            object rsltswfolders = null;
            object softwarelicense = null;
            object objtype = null;
            GetReferenceLicenseCount_retVal = (Int16)0;

            rsltswfolders = "";
            softwarelicense = VBScriptConstants.Nothing;
            objtype = "";

            //Prüfen ob es Software Lizenzobjekte/Lizenzumschläge unterhalb des Folders gibt.
            object byrefalias29 = hlasc_softwarelicensefolderview;
            try
            {
                rsltswfolders = _.VAL(_.CALL(this, hlswfolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias29, v56 => { byrefalias29 = v56; })));
            }
            finally { hlasc_softwarelicensefolderview = byrefalias29; }

            var enumerationContent2 = _.ENUMERABLE(rsltswfolders).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                softwarelicense = enumerationContent2.Current;
                objtype = _.VAL(_.CALL(this, softwarelicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objtype), "LicenseFolder")))
                {
                    object byrefalias30 = hlcontext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias30, v57 => { byrefalias30 = v57; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlcontext = byrefalias30; }
                    if (_.IF(_.GT(_.NullableNUM(GetReferenceLicenseCount_retVal), (Int16)0)))
                    {
                        return GetReferenceLicenseCount_retVal;
                    }
                }
                if (_.IF(_.AND(_.EQ(_.NullableSTR(objtype), "SoftwareLicense"), _.EQ(_.CBOOL(chkfolderonly), false))))
                {
                    object byrefalias31 = hlcontext;
                    try
                    {
                        GetReferenceLicenseCount_retVal = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias31, v58 => { byrefalias31 = v58; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlcontext = byrefalias31; }
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
        public object checkforsoftwaresuitefolder(ref object hlcontext, ref object hlparentswfolder, ref object pdict, ref object hlasc_softwarelicensefolderview)
        {
            object CheckForSoftwareSuiteFolder_retVal = null;
            object retval = null;
            object nextswfolder = null;
            object checksoftwaresuite = null;
            CheckForSoftwareSuiteFolder_retVal = "";
            retval = (Int16)0;
            nextswfolder = "";

            //Festhalten auf welcher Ebene ggf. eine Software Suite oberhalb des
            //Start Folders existiert. Die Variable muss von außen mit einem Startwert
            //initialisiert werden.
            if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALL(this, pdict, _.ARGS.Val("SoftwareSuiteFolderLevel"))), (Int16)0), _.EQ(_.NullableSTR(_.CALL(this, pdict, _.ARGS.Val("SoftwareSuiteFolderLevel"))), ""))))
            {
                _.SET((Int16)1, this, pdict, null, _.ARGS.Val("SoftwareSuiteFolderLevel"));
            }
            else
            {
                _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SoftwareSuiteFolderLevel")), (Int16)1), this, pdict, null, _.ARGS.Val("SoftwareSuiteFolderLevel"));
            }

            //Amhand des Flags "Software Suite" festellen ob ein Lizenzumschlag als Software Suite
            //gekennzeichnet ist. Falls Ja, Name des Umschlags auslesen und Funktion abbrechen.
            checksoftwaresuite = false;
            object byrefalias32 = hlcontext, byrefalias33 = hlparentswfolder;
            try
            {
                checksoftwaresuite = _.VAL(_.CALL(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias32, v59 => { byrefalias32 = v59; }).Ref(byrefalias33, v60 => { byrefalias33 = v60; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlcontext = byrefalias32; hlparentswfolder = byrefalias33; }
            if (_.IF(_.EQ(_.CBOOL(checksoftwaresuite), true)))
            {
                _.SET(_.VAL(_.CALL(this, hlparentswfolder, "GetValue", _.ARGS.Val("OrganizationGeneral.Name").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))), this, pdict, null, _.ARGS.Val("SoftwareSuiteFolder"));
                return CheckForSoftwareSuiteFolder_retVal;
            }

            //Wenn sich mindestens noch ein weiterer Lizenzumschlag oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            object byrefalias34 = hlasc_softwarelicensefolderview;
            try
            {
                nextswfolder = _.VAL(_.CALL(this, hlparentswfolder, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias34, v61 => { byrefalias34 = v61; })));
            }
            finally { hlasc_softwarelicensefolderview = byrefalias34; }
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(nextswfolder)), (Int16)0)))
            {
                object byrefalias35 = hlcontext, byrefalias36 = pdict, byrefalias37 = hlasc_softwarelicensefolderview;
                try
                {
                    retval = _.VAL(_.CALL(this, _outer, "CheckForSoftwareSuiteFolder", _.ARGS.Ref(byrefalias35, v62 => { byrefalias35 = v62; }).RefIfArray(nextswfolder, _.ARGS.Val((Int16)0)).Ref(byrefalias36, v63 => { byrefalias36 = v63; }).Ref(byrefalias37, v64 => { byrefalias37 = v64; })));
                }
                finally { hlcontext = byrefalias35; pdict = byrefalias36; hlasc_softwarelicensefolderview = byrefalias37; }
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
        public object setlicensecounter(ref object hlcontext, ref object hlswfolder, ref object pdict, ref object assocname)
        {
            object SetLicenseCounter_retVal = null;
            object retval = null;
            object checksoftwaresuite = null;
            object checkliccontrbyserver = null;
            object nextswfolder = null;
            object a = null;
            SetLicenseCounter_retVal = (Int16)0;
            retval = (Int16)0;

            //Dictionary Einträge initalisieren
            _.SET("", this, pdict, null, _.ARGS.Val("SoftwareLicenses"));
            _.SET((Int16)0, this, pdict, null, _.ARGS.Val("SumRefLicCounter"));
            _.SET((Int16)0, this, pdict, null, _.ARGS.Val("SumInstLicCounter"));
            _.SET((Int16)0, this, pdict, null, _.ARGS.Val("SumFreeLicCounter"));

            //Prüfen ob es Software Lizenzobjekte unterhalb des Folders gibt.
            object byrefalias38 = assocname;
            try
            {
                _.SET(_.VAL(_.CALL(this, hlswfolder, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Ref(byrefalias38, v65 => { byrefalias38 = v65; }))), this, pdict, null, _.ARGS.Val("SoftwareLicenses"));
            }
            finally { assocname = byrefalias38; }

            //Amhand des Flags "Software Suite" entscheiden ob alle Objekte oder nur Folder
            //Objekte gezählt werden müssen
            checksoftwaresuite = false;
            object byrefalias39 = hlcontext, byrefalias40 = hlswfolder;
            try
            {
                checksoftwaresuite = _.VAL(_.CALL(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias39, v66 => { byrefalias39 = v66; }).Ref(byrefalias40, v67 => { byrefalias40 = v67; }).Val("SoftwareLicenseFolderDetail.FlagSoftwareSuite").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlcontext = byrefalias39; hlswfolder = byrefalias40; }

            bool ifResult5;
            object byrefalias41 = pdict;
            try
            {
                ifResult5 = _.IF(_.GTE(_.NullableNUM(_.UBOUND(_.CALL(this, byrefalias41, _.ARGS.Val("SoftwareLicenses")))), (Int16)0));
            }
            finally { pdict = byrefalias41; }
            if (ifResult5)
            {
                if (_.IF(_.EQ(_.CBOOL(checksoftwaresuite), false)))
                {
                    object byrefalias42 = hlcontext, byrefalias43 = pdict;
                    try
                    {
                        _.CALL(this, _outer, "CalcAllLicCounter", _.ARGS.Ref(byrefalias42, v68 => { byrefalias42 = v68; }).Ref(byrefalias43, v69 => { byrefalias43 = v69; }));
                    }
                    finally { hlcontext = byrefalias42; pdict = byrefalias43; }
                }
                else
                {
                    object byrefalias44 = hlcontext, byrefalias45 = pdict;
                    try
                    {
                        _.CALL(this, _outer, "CalcFolderLicCounter", _.ARGS.Ref(byrefalias44, v70 => { byrefalias44 = v70; }).Ref(byrefalias45, v71 => { byrefalias45 = v71; }));
                    }
                    finally { hlcontext = byrefalias44; pdict = byrefalias45; }
                }
            }
            //Gesatmzahl der Lizenzen in den Lizenzumschlag zurückschreiben
            object byrefalias46 = pdict;
            try
            {
                _.CALL(this, hlswfolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias46, _.ARGS.Val("SumRefLicCounter")));
            }
            finally { pdict = byrefalias46; }
            object byrefalias47 = pdict;
            try
            {
                _.CALL(this, hlswfolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias47, _.ARGS.Val("SumInstLicCounter")));
            }
            finally { pdict = byrefalias47; }

            //Wenn die Lizenzkontrolle durch den Applikations Server erfolgt ("Lizenzkontrolle durch Server")
            //dann die Anzahl freier Lizenzen immer auf den Wert "0" setzen.
            checkliccontrbyserver = false;
            object byrefalias48 = hlcontext, byrefalias49 = hlswfolder;
            try
            {
                checkliccontrbyserver = _.VAL(_.CALL(this, _outer, "GetFlagValue", _.ARGS.Ref(byrefalias48, v72 => { byrefalias48 = v72; }).Ref(byrefalias49, v73 => { byrefalias49 = v73; }).Val("SoftwareLicenseFolderDetail.FlagLicenseControlledByServer").Val((Int16)0).Val((Int16)0)));
            }
            finally { hlcontext = byrefalias48; hlswfolder = byrefalias49; }
            if (_.IF(_.EQ(_.CBOOL(checkliccontrbyserver), true)))
            {
                _.SET((Int16)0, this, pdict, null, _.ARGS.Val("SumFreeLicCounter"));
            }
            object byrefalias50 = pdict;
            try
            {
                _.CALL(this, hlswfolder, "SetValue", _.ARGS.Val("SoftwareLicenseCounter.FreeLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).RefIfArray(byrefalias50, _.ARGS.Val("SumFreeLicCounter")));
            }
            finally { pdict = byrefalias50; }

            //Erst prüfen, um welchen OU Typ es sich handelt, damit die richtige Assoziationsdefinition
            //für die nächste Abfrage gewählt werden kann.
            nextswfolder = "";
            a = "";
            a = _.VAL(_.CALL(this, hlswfolder, "GetType"));
            if (_.IF(_.EQ(_.NullableSTR(a), "LicenseFolder")))
            {
                object byrefalias51 = assocname;
                try
                {
                    nextswfolder = _.VAL(_.CALL(this, hlswfolder, "GetItems", _.ARGS.Val(65536).Val((Int16)0).Val((Int16)0).Ref(byrefalias51, v74 => { byrefalias51 = v74; })));
                }
                finally { assocname = byrefalias51; }
            }
            //Wenn sich mindestens noch eine weitere OU oberhalb der aktuellen befindet,
            //dann wird die Funktion erneut aufgerufen. Anderfalls wird die Function beendet.
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(nextswfolder)), (Int16)0)))
            {
                object byrefalias52 = hlcontext, byrefalias53 = pdict, byrefalias54 = assocname;
                try
                {
                    retval = _.VAL(_.CALL(this, _outer, "SetLicenseCounter", _.ARGS.Ref(byrefalias52, v75 => { byrefalias52 = v75; }).RefIfArray(nextswfolder, _.ARGS.Val((Int16)0)).Ref(byrefalias53, v76 => { byrefalias53 = v76; }).Ref(byrefalias54, v77 => { byrefalias54 = v77; })));
                }
                finally { hlcontext = byrefalias52; pdict = byrefalias53; assocname = byrefalias54; }
            }
            else
            {
                return SetLicenseCounter_retVal;
            }
            return SetLicenseCounter_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        public object isvalidobject(ref object obj)
        {
            return _.VAL(_.AND(_.ISOBJECT(obj), _.NOT(_.IS(obj, VBScriptConstants.Nothing))));
        }

        //----------------------------------------------------------------------------------------------------------
        public void calcallliccounter(ref object hlcontext, ref object pdict)
        {
            object swrefliccounter = null;
            object swinstcounter = null;
            object softwarelicense = null;
            object objtype = null;
            object lstlicstatus = null;
            swrefliccounter = (Int16)0;
            swinstcounter = (Int16)0;
            softwarelicense = VBScriptConstants.Nothing;
            objtype = "";
            lstlicstatus = "";

            var enumerationContent3 = _.ENUMERABLE(_.CALL(this, pdict, _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                softwarelicense = enumerationContent3.Current;
                objtype = _.VAL(_.CALL(this, softwarelicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.EQ(_.NullableSTR(objtype), "SoftwareLicense")))
                {
                    lstlicstatus = _.VAL(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseDetail.LicenseStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(lstlicstatus), "LicenseStatusValid")))
                    {
                        object byrefalias55 = hlcontext;
                        try
                        {
                            swrefliccounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias55, v78 => { byrefalias55 = v78; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlcontext = byrefalias55; }
                        _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), swrefliccounter), this, pdict, null, _.ARGS.Val("SumRefLicCounter"));
                    }
                }
                else
                {
                    if (_.IF(_.OR(_.EQ(_.NullableSTR(objtype), "LicenseFolder"), _.EQ(_.NullableSTR(objtype), "Software"))))
                    {
                        object byrefalias56 = hlcontext;
                        try
                        {
                            swrefliccounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias56, v79 => { byrefalias56 = v79; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlcontext = byrefalias56; }
                        _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), swrefliccounter), this, pdict, null, _.ARGS.Val("SumRefLicCounter"));
                        object byrefalias57 = hlcontext;
                        try
                        {
                            swinstcounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias57, v80 => { byrefalias57 = v80; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlcontext = byrefalias57; }
                        _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SumInstLicCounter")), swinstcounter), this, pdict, null, _.ARGS.Val("SumInstLicCounter"));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SET(_.SUBT(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), _.CALL(this, pdict, _.ARGS.Val("SumInstLicCounter"))), this, pdict, null, _.ARGS.Val("SumFreeLicCounter"));

        }

        //----------------------------------------------------------------------------------------------------------
        public void calcfolderliccounter(ref object hlcontext, ref object pdict)
        {
            object swrefliccounter = null;
            object swinstcounter = null;
            object softwarelicense = null;
            object objtype = null;
            object lstlicstatus = null;

            swrefliccounter = (Int16)0;
            swinstcounter = (Int16)0;
            softwarelicense = VBScriptConstants.Nothing;
            objtype = "";
            lstlicstatus = "";

            var enumerationContent4 = _.ENUMERABLE(_.CALL(this, pdict, _.ARGS.Val("SoftwareLicenses"))).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                softwarelicense = enumerationContent4.Current;
                objtype = _.VAL(_.CALL(this, softwarelicense, "GetType", _.ARGS.ForceBrackets()));
                if (_.IF(_.OR(_.EQ(_.NullableSTR(objtype), "LicenseFolder"), _.EQ(_.NullableSTR(objtype), "Software"))))
                {
                    object byrefalias58 = hlcontext;
                    try
                    {
                        swrefliccounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias58, v81 => { byrefalias58 = v81; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlcontext = byrefalias58; }
                    _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), swrefliccounter), this, pdict, null, _.ARGS.Val("SumRefLicCounter"));

                    object byrefalias59 = hlcontext;
                    try
                    {
                        swinstcounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias59, v82 => { byrefalias59 = v82; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.InstalledLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                    }
                    finally { hlcontext = byrefalias59; }
                    if (_.IF(_.GT(swinstcounter, _.CALL(this, pdict, _.ARGS.Val("SumInstLicCounter")))))
                    {
                        _.SET(_.VAL(swinstcounter), this, pdict, null, _.ARGS.Val("SumInstLicCounter"));
                    }
                }
                if (_.IF(_.EQ(_.NullableSTR(objtype), "SoftwareLicense")))
                {
                    lstlicstatus = _.VAL(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseDetail.LicenseStatus").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    if (_.IF(_.EQ(_.NullableSTR(lstlicstatus), "LicenseStatusValid")))
                    {
                        object byrefalias60 = hlcontext;
                        try
                        {
                            swrefliccounter = _.VAL(_.CALL(this, _outer, "CheckIntegerValue", _.ARGS.Ref(byrefalias60, v83 => { byrefalias60 = v83; }).Val(_.CALL(this, softwarelicense, "GetValue", _.ARGS.Val("SoftwareLicenseCounter.ReferenceLicenseCount").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)))));
                        }
                        finally { hlcontext = byrefalias60; }
                        _.SET(_.ADD(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), swrefliccounter), this, pdict, null, _.ARGS.Val("SumRefLicCounter"));
                    }
                }
            }
            //Anzahl freier Lizenzen berechnen und in den Folder schreiben.
            _.SET(_.SUBT(_.CALL(this, pdict, _.ARGS.Val("SumRefLicCounter")), _.CALL(this, pdict, _.ARGS.Val("SumInstLicCounter"))), this, pdict, null, _.ARGS.Val("SumFreeLicCounter"));
        }

        //----------------------------------------------------------------------------------------------------------
        //Diese Function überprüft den ganzzahligen Wert (Integer).
        public object checkintegervalue(ref object hlcontext, ref object intval)
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
        public object oncreate_hasassociationtodelete(ref object hlcontext, ref object ascdefname, ref object hlobjb)
        {
            object OnCreate_HasAssociationToDelete_retVal = null;
            object result = null;
            object cassociationchanges = null;
            object oassociationchange = null;
            object ascdefnamechange = null;
            object ixac = null;
            result = false;
            cassociationchanges = (Int16)0;
            cassociationchanges = _.VAL(_.CALL(this, hlcontext, "GetAssociationChangesCount"));

            oassociationchange = VBScriptConstants.Nothing;
            ascdefnamechange = "";
            ixac = (Int16)0;

            var loopEnd = _.NUM(_.SUBT(cassociationchanges, (Int16)1));
            var loopStart = _.NUM((Int16)0, loopEnd, (Int16)1);
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (ixac = loopStart; _.StrictLTE(ixac, loopEnd); ixac = _.ADD(ixac, (Int16)1))
                {
                    oassociationchange = _.OBJ(_.CALL(this, hlcontext, "GetAssociationChangeAt", _.ARGS.Ref(ixac, v84 => { ixac = v84; })));

                    ascdefnamechange = _.VAL(_.CALL(this, oassociationchange, "AssociationType"));

                    if (_.IF(_.CALL(this, oassociationchange, "IsToDelete")))
                    {
                        if (_.IF(_.EQ(ascdefnamechange, ascdefname)))
                        {
                            if (_.IF(_.EQ(_.CALL(this, hlobjb, "GetID"), _.CALL(this, oassociationchange, "EndB", "GetID"))))
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
        public object oncreate_hasassociationtocreate(ref object hlcontext, ref object ascdefname, ref object hlobjb)
        {
            object OnCreate_HasAssociationToCreate_retVal = null;
            object result = null;
            object cassociationchanges = null;
            object oassociationchange = null;
            object ascdefnamechange = null;
            object ixac = null;
            result = false;
            cassociationchanges = (Int16)0;
            cassociationchanges = _.VAL(_.CALL(this, hlcontext, "GetAssociationChangesCount"));

            oassociationchange = VBScriptConstants.Nothing;
            ascdefnamechange = "";
            ixac = (Int16)0;

            var loopEnd2 = _.NUM(_.SUBT(cassociationchanges, (Int16)1));
            var loopStart2 = _.NUM((Int16)0, loopEnd2, (Int16)1);
            if (_.StrictLTE(loopStart2, loopEnd2))
            {
                for (ixac = loopStart2; _.StrictLTE(ixac, loopEnd2); ixac = _.ADD(ixac, (Int16)1))
                {
                    oassociationchange = _.OBJ(_.CALL(this, hlcontext, "GetAssociationChangeAt", _.ARGS.Ref(ixac, v85 => { ixac = v85; })));

                    ascdefnamechange = _.VAL(_.CALL(this, oassociationchange, "AssociationType"));

                    if (_.IF(_.CALL(this, oassociationchange, "IsToCreate")))
                    {
                        if (_.IF(_.EQ(ascdefnamechange, ascdefname)))
                        {
                            if (_.IF(_.EQ(_.CALL(this, hlobjb, "GetID"), _.CALL(this, oassociationchange, "EndB", "GetID"))))
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

        public object ondelete_hasassociationtocreate(ref object hlcontext, ref object ascdefname, ref object hlobjb)
        {
            object OnDelete_HasAssociationToCreate_retVal = null;
            object result = null;
            object cassociationchanges = null;
            object oassociationchange = null;
            object ascdefnamechange = null;
            object ixac = null;
            // bool
            result = false;

            //Anzahl der zu erstellenden oder löschenden Assoziationen
            cassociationchanges = (Int16)0;
            cassociationchanges = _.VAL(_.CALL(this, hlcontext, "GetAssociationChangesCount"));

            oassociationchange = VBScriptConstants.Nothing;
            ascdefnamechange = "";
            ixac = (Int16)0;

            var loopEnd3 = _.NUM(_.SUBT(cassociationchanges, (Int16)1));
            var loopStart3 = _.NUM((Int16)0, loopEnd3, (Int16)1);
            if (_.StrictLTE(loopStart3, loopEnd3))
            {
                for (ixac = loopStart3; _.StrictLTE(ixac, loopEnd3); ixac = _.ADD(ixac, (Int16)1))
                {

                    //Für jede Assoziations Änderung wird das entsprechende Infos (Objekt    ) ausgelsen.
                    oassociationchange = _.OBJ(_.CALL(this, hlcontext, "GetAssociationChangeAt", _.ARGS.Ref(ixac, v86 => { ixac = v86; })));
                    //Def Name der Assoc ermitteln, die angelegt werden soll
                    ascdefnamechange = _.VAL(_.CALL(this, oassociationchange, "AssociationType"));

                    if (_.IF(_.CALL(this, oassociationchange, "IsToCreate")))
                    {
                        //Überprüfen ob die gewünschte Assoc auch angelegt werden soll.
                        if (_.IF(_.EQ(ascdefnamechange, ascdefname)))
                        {
                            if (_.IF(_.EQ(_.CALL(this, hlobjb, "GetID"), _.CALL(this, oassociationchange, "EndB", "GetID"))))
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
        public object getassociatedorganizationalunit(ref object hlcontext, ref object lcid, ref object hlchild, ref object pdict, ref object outparentdefname)
        {
            object GetAssociatedOrganizationalUnit_retVal = null;
            object rsltparent = null;
            object objparent = null;
            GetAssociatedOrganizationalUnit_retVal = "";
            outparentdefname = "";

            rsltparent = "";
            object byrefalias61 = pdict;
            try
            {
                rsltparent = _.VAL(_.CALL(this, hlchild, "GetItems", _.ARGS.Val(65536).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).RefIfArray(byrefalias61, _.ARGS.Val("AssocID"))));
            }
            finally { pdict = byrefalias61; }
            if (_.IF(_.GTE(_.UBOUND(rsltparent), _.CALL(this, pdict, _.ARGS.Val("ParentCounter")))))
            {
                objparent = VBScriptConstants.Nothing;
                var enumerationContent5 = _.ENUMERABLE(rsltparent).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent5.MoveNext())
                        break;
                    objparent = enumerationContent5.Current;
                    object byrefalias62 = pdict;
                    try
                    {
                        GetAssociatedOrganizationalUnit_retVal = _.VAL(_.CALL(this, objparent, "GetValue", _.ARGS.RefIfArray(byrefalias62, _.ARGS.Val("AttrName")).Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    }
                    finally { pdict = byrefalias62; }
                    object byrefalias63 = lcid;
                    try
                    {
                        outparentdefname = _.VAL(_.CALL(this, hlcontext, "GetDisplayName", _.ARGS.Val(_.CALL(this, objparent, "GetValue", _.ARGS.Val("HLOBJECTINFO.DEFID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))).Ref(byrefalias63, v87 => { byrefalias63 = v87; })));
                    }
                    finally { lcid = byrefalias63; }
                    break;
                }
            }
            return GetAssociatedOrganizationalUnit_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------
        public object mig_createxmldocument(ref object hlsrvcontext, ref object pdict)
        {
            object MIG_CreateXMLDocument_retVal = null;
            object objxmldoc = null;
            object xmlproinc = null;
            object xmlroot = null;
            object nodesession = null;

            //XML-Objekt erstellen
            objxmldoc = VBScriptConstants.Nothing;
            objxmldoc = _.OBJ(_.CREATEOBJECT("Msxml2.DOMDocument"));

            //XML-Processing Instruction hinzufügen
            xmlproinc = VBScriptConstants.Nothing;
            xmlproinc = _.OBJ(_.CALL(this, objxmldoc, "createProcessingInstruction", _.ARGS.Val("xml").Val("version='1.0' encoding='UTF-8'")));
            _.CALL(this, objxmldoc, "insertBefore", _.ARGS.Ref(xmlproinc, v88 => { xmlproinc = v88; }).Val(_.CALL(this, objxmldoc, "firstChild")));

            //Root-Element erstellen
            xmlroot = _.OBJ(_.CALL(this, objxmldoc, "CreateElement", _.ARGS.Val("ASAPBatch")));
            _.CALL(this, objxmldoc, "AppendChild", _.ARGS.Val(xmlroot));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("xmlns").Val("http://www.brainware.ch/operationsmanager/asap-batch/1.1"));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("xmlns:dt").Val("http://www.brainware.ch/operationsmanager/wf/changemanagement/columbus/datatypes/1.1"));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("xmlns:xsi").Val("http://www.w3.org/2001/XMLSchema-instance"));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("xsi:schemaLocation").Val("http://www.brainware.ch/operationsmanager/asap-batch/1.1 asap-batch-1.1.xsd"));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("version").Val("1.1"));
            _.CALL(this, xmlroot, "SetAttribute", _.ARGS.Val("responseRequired").Val("Yes"));

            //Das Node Session hinzufügen
            nodesession = _.OBJ(_.CALL(this, objxmldoc, "CreateElement", _.ARGS.Val("Session")));
            _.CALL(this, xmlroot, "AppendChild", _.ARGS.Val(nodesession));
            _.CALL(this, nodesession, "SetAttribute", _.ARGS.Val("id").Val("s1"));
            _.CALL(this, nodesession, "SetAttribute", _.ARGS.Val("loginname").Val("foreignSystems\\assetcolumbus"));
            _.CALL(this, nodesession, "SetAttribute", _.ARGS.Val("password").Val(""));

            //XML Dokument inkl. Header an das Dictionary übergeben
            _.SET(_.OBJ(objxmldoc), this, pdict, null, _.ARGS.Val("XMLDocument"));
            return MIG_CreateXMLDocument_retVal;
        }

        //---------------------------------------------------------------------
        public object mig_createaddxml2columbus(ref object hlsrvcontext, ref object pdict)
        {
            object MIG_CreateADDXML2Columbus_retVal = null;
            object xmlroot = null;
            object nodecreateinstancerq = null;
            object nodeobserverkey = null;
            object nodecontextdata = null;
            object nodeadddeviceactualparams = null;
            object nodedeviceidentification = null;
            object nodedevicename = null;
            object nodecmpyname = null;
            object nodedomain = null;
            object nodecostcenter = null;
            object nodemacaddress = null;
            object nodesubnetmask = null;
            object nodehwtype = null;
            object nodeostype = null;
            object nodeactstate = null;

            //Root Element aus dem XML ermitteln.
            xmlroot = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodecreateinstancerq = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("CreateInstanceRq")));
            _.CALL(this, xmlroot, "AppendChild", _.ARGS.Val(nodecreateinstancerq));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("id").Val("e7"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfpNs").Val("ch.bw.wf.changemgmt.columbus_adddevice"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfmNs").Val("Columbus Changemanagement"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("sessionId").Val("s1"));

            //Das Node ObserverKey hinzufügen
            nodeobserverkey = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ObserverKey")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodeobserverkey));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("ObserverKey"))), this, nodeobserverkey, "Text");

            //Das Container Node ContextData hinzufügen
            nodecontextdata = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ContextData")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodecontextdata));

            //Das Container Node AddDeviceActualParams hinzufügen
            nodeadddeviceactualparams = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:AddDeviceActualParams")));
            _.CALL(this, nodecontextdata, "AppendChild", _.ARGS.Val(nodeadddeviceactualparams));

            //Das Container Node DeviceIdentification hinzufügen
            nodedeviceidentification = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceIdentification")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodedeviceidentification));

            //Das Node DeviceName hinzufügen
            nodedevicename = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceName")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedevicename));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("DeviceName"))), this, nodedevicename, "Text");

            //Das Node CompanyName hinzufügen
            nodecmpyname = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:CompanyName")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodecmpyname));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("CompanyName"))), this, nodecmpyname, "Text");

            //Das Node Domain hinzufügen
            nodedomain = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:Domain")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedomain));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("Domain"))), this, nodedomain, "Text");

            //Das Node CostCenter hinzufügen
            nodecostcenter = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:CostCenter")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodecostcenter));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("CostCenter"))), this, nodecostcenter, "Text");

            //Das Node MACAdess hinzufügen
            nodemacaddress = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:MACAddress")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodemacaddress));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("MACAddress"))), this, nodemacaddress, "Text");

            //Das Node SubnetMask hinzufügen
            nodesubnetmask = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:SubnetMask")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodesubnetmask));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("SubnetMask"))), this, nodesubnetmask, "Text");

            //Das Node HwTypeId hinzufügen
            nodehwtype = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:HwTypeId")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodehwtype));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("HwTypeId"))), this, nodehwtype, "Text");

            //Das Node OsTypeId hinzufügen
            nodeostype = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:OsTypeId")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodeostype));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("OsTypeId"))), this, nodeostype, "Text");

            //Das Node ActivationState hinzufügen
            nodeactstate = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:ActivationState")));
            _.CALL(this, nodeadddeviceactualparams, "AppendChild", _.ARGS.Val(nodeactstate));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("ActivationState"))), this, nodeactstate, "Text");

            return MIG_CreateADDXML2Columbus_retVal;
        }

        //---------------------------------------------------------------------
        public object mig_createchgxml2columbus(ref object hlsrvcontext, ref object pdict)
        {
            object MIG_CreateCHGXML2Columbus_retVal = null;
            object xmlroot = null;
            object nodecreateinstancerq = null;
            object nodeobserverkey = null;
            object nodecontextdata = null;
            object nodechgdeviceactualparams = null;
            object nodedeviceidentification = null;
            object nodedevicename = null;
            object nodedomain = null;
            object nodecmpyname = null;
            object nodecostcenter = null;
            object nodemacaddress = null;
            object nodesubnetmask = null;
            object nodehwtype = null;
            object nodeostype = null;
            object nodeactstate = null;

            //Root Element aus dem XML ermitteln.
            xmlroot = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodecreateinstancerq = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("CreateInstanceRq")));
            _.CALL(this, xmlroot, "AppendChild", _.ARGS.Val(nodecreateinstancerq));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("id").Val("e7"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfpNs").Val("ch.bw.wf.changemgmt.columbus_chgdevice"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfmNs").Val("Columbus Changemanagement"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("sessionId").Val("s1"));

            //Das Node ObserverKey hinzufügen
            nodeobserverkey = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ObserverKey")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodeobserverkey));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("ObserverKey"))), this, nodeobserverkey, "Text");

            //Das Container Node ContextData hinzufügen
            nodecontextdata = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ContextData")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodecontextdata));

            //Das Container Node AddDeviceActualParams hinzufügen
            nodechgdeviceactualparams = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:ChangeDeviceActualParams")));
            _.CALL(this, nodecontextdata, "AppendChild", _.ARGS.Val(nodechgdeviceactualparams));

            //Das Container Node DeviceIdentification hinzufügen
            nodedeviceidentification = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceIdentification")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodedeviceidentification));

            //Das Node DeviceName hinzufügen
            nodedevicename = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceName")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedevicename));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("DeviceName"))), this, nodedevicename, "Text");

            //Das Node Domain hinzufügen
            nodedomain = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:Domain")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedomain));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("Domain"))), this, nodedomain, "Text");

            //Das Node CompanyName hinzufügen
            nodecmpyname = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:CompanyName")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodecmpyname));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("CompanyName"))), this, nodecmpyname, "Text");

            //Das Node CostCenter hinzufügen
            nodecostcenter = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:CostCenter")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodecostcenter));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("CostCenter"))), this, nodecostcenter, "Text");

            //Das Node MACAdess hinzufügen
            nodemacaddress = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:MACAddress")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodemacaddress));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("MACAddress"))), this, nodemacaddress, "Text");

            //Das Node SubnetMask hinzufügen
            nodesubnetmask = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:SubnetMask")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodesubnetmask));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("SubnetMask"))), this, nodesubnetmask, "Text");

            //Das Node HwTypeId hinzufügen
            nodehwtype = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:HwTypeId")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodehwtype));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("HwTypeId"))), this, nodehwtype, "Text");

            //Das Node OsTypeId hinzufügen
            nodeostype = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:OsTypeId")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodeostype));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("OsTypeId"))), this, nodeostype, "Text");

            //Das Node ActivationState hinzufügen
            nodeactstate = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:ActivationState")));
            _.CALL(this, nodechgdeviceactualparams, "AppendChild", _.ARGS.Val(nodeactstate));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("ActivationState"))), this, nodeactstate, "Text");

            return MIG_CreateCHGXML2Columbus_retVal;
        }

        //---------------------------------------------------------------------
        public object mig_createdelxml2columbus(ref object hlsrvcontext, ref object pdict)
        {
            object MIG_CreateDELXML2Columbus_retVal = null;
            object xmlroot = null;
            object nodecreateinstancerq = null;
            object nodeobserverkey = null;
            object nodecontextdata = null;
            object noderemovedeviceactualparams = null;
            object nodedeviceidentification = null;
            object nodedevicename = null;
            object nodedomain = null;

            //Root Element aus dem XML ermitteln.
            xmlroot = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "DocumentElement"));

            //Das Node CreateInstanceReq hinzufügen
            nodecreateinstancerq = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("CreateInstanceRq")));
            _.CALL(this, xmlroot, "AppendChild", _.ARGS.Val(nodecreateinstancerq));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("id").Val("e7"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfpNs").Val("ch.bw.wf.changemgmt.columbus_removedevice"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("wfmNs").Val("Columbus Changemanagement"));
            _.CALL(this, nodecreateinstancerq, "SetAttribute", _.ARGS.Val("sessionId").Val("s1"));

            //Das Node ObserverKey hinzufügen
            nodeobserverkey = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ObserverKey")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodeobserverkey));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("ObserverKey"))), this, nodeobserverkey, "Text");

            //Das Container Node ContextData hinzufügen
            nodecontextdata = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("ContextData")));
            _.CALL(this, nodecreateinstancerq, "AppendChild", _.ARGS.Val(nodecontextdata));

            //Das Container Node AddDeviceActualParams hinzufügen
            noderemovedeviceactualparams = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:RemoveDeviceActualParams")));
            _.CALL(this, nodecontextdata, "AppendChild", _.ARGS.Val(noderemovedeviceactualparams));

            //Das Container Node DeviceIdentification hinzufügen
            nodedeviceidentification = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceIdentification")));
            _.CALL(this, noderemovedeviceactualparams, "AppendChild", _.ARGS.Val(nodedeviceidentification));

            //Das Node DeviceName hinzufügen
            nodedevicename = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:DeviceName")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedevicename));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("DeviceName"))), this, nodedevicename, "Text");

            //Das Node CompanyName hinzufügen
            //Dim nodeCmpyName : Set nodeCmpyName = pDict("XMLDocument").CreateElement("dt:CompanyName")
            //nodeDeviceIdentification.AppendChild (nodeCmpyName)
            //nodeCmpyName.Text = pDict("CompanyName")

            //Das Node Domain hinzufügen
            nodedomain = _.OBJ(_.CALL(this, _.CALL(this, pdict, _.ARGS.Val("XMLDocument")), "CreateElement", _.ARGS.Val("dt:Domain")));
            _.CALL(this, nodedeviceidentification, "AppendChild", _.ARGS.Val(nodedomain));
            _.SET(_.VAL(_.CALL(this, pdict, _.ARGS.Val("Domain"))), this, nodedomain, "Text");

            return MIG_CreateDELXML2Columbus_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //Wenn beide Werte ein Datum sind, muss geprüft werden ob das Enddatum nach dem
        //Start Datum liegt. Falls nicht wird "False" zurückgegeben.
        public object migcheckdateperiod(ref object hlcontext, ref object startdate, ref object enddate)
        {
            object MigCheckDatePeriod_retVal = null;
            MigCheckDatePeriod_retVal = false;

            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _, "DATEPART", _.ARGS.Val("d").Val(_.CDATE(startdate)))), "0")))
            {
                if (_.IF(_.LT(_.CALL(this, _, "DATEPART", _.ARGS.Val("d").Val(_.CDATE(startdate))), _.CALL(this, _, "DATEPART", _.ARGS.Val("d").Val(_.CDATE(enddate))))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    MigCheckDatePeriod_retVal = true;
                }

                if (_.IF(_.GT(_.CALL(this, _, "DATEPART", _.ARGS.Val("yyyy").Val(_.CDATE(startdate))), _.CALL(this, _, "DATEPART", _.ARGS.Val("yyyy").Val(_.CDATE(enddate))))))
                {
                    MigCheckDatePeriod_retVal = false;
                }
                else
                {
                    if (_.IF(_.GT(_.CALL(this, _, "DATEPART", _.ARGS.Val("y").Val(_.CDATE(startdate))), _.CALL(this, _, "DATEPART", _.ARGS.Val("y").Val(_.CDATE(enddate))))))
                    {
                        if (_.IF(_.LT(_.CALL(this, _, "DATEPART", _.ARGS.Val("yyyy").Val(_.CDATE(startdate))), _.CALL(this, _, "DATEPART", _.ARGS.Val("yyyy").Val(_.CDATE(enddate))))))
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
        public object mig_checkcostcenter(ref object hlsrvcontext, ref object strcostcenter)
        {
            object MIG_CheckCostCenter_retVal = null;
            object srchquery = null;
            object qry = null;
            object rsltquery = null;
            MIG_CheckCostCenter_retVal = false;

            srchquery = "";
            srchquery = _.CONCAT("SEARCH Division WHERE OrganizationBilling.CostCenter_CA.CostCenter = \"", strcostcenter, "\"");
            qry = VBScriptConstants.Nothing;
            rsltquery = "";
            qry = _.OBJ(_.CALL(this, hlsrvcontext, "OpenSearch", _.ARGS.Ref(srchquery, v89 => { srchquery = v89; })));
            rsltquery = _.VAL(_.CALL(this, qry, "GetItems", _.ARGS.Val((Int16)0).Val(_.SUBT((Int16)1)).Val(_.SUBT((Int16)1)).Val((Int16)0)));
            if (_.IF(_.GTE(_.NullableNUM(_.UBOUND(rsltquery)), (Int16)0)))
            {
                MIG_CheckCostCenter_retVal = true;
            }

            return MIG_CheckCostCenter_retVal;
        }

        public object checkagenthasmigpartnerid(ref object hlcontext, ref object relobjmigpartnerid)
        {
            object CheckAgentHasMIGPartnerID_retVal = null;
            object flagauthorized = null;
            object intagentid = null;
            object objperson = null;
            object strpersoninternalmigpartnerids = null;
            //BOOL

            flagauthorized = false;
            intagentid = _.VAL(_.CALL(this, hlcontext, "GetAgentID", _.ARGS.ForceBrackets()));
            objperson = VBScriptConstants.Nothing;

            objperson = _.OBJ(_.CALL(this, hlcontext, "GetPersonOfAgent", _.ARGS.Ref(intagentid, v90 => { intagentid = v90; })));

            bool ifResult6;
            object byrefalias64 = hlcontext;
            try
            {
                ifResult6 = _.IF(_.EQ(_.CALL(this, _outer, "IsHLObject", _.ARGS.Ref(byrefalias64, v93 => { byrefalias64 = v93; }).Ref(objperson, v94 => { objperson = v94; })), true));
            }
            finally { hlcontext = byrefalias64; }
            if (ifResult6)
            {

                if (_.IF(_.NOTEQ(_.NullableSTR(relobjmigpartnerid), "")))
                {

                    strpersoninternalmigpartnerids = _.VAL(_.CALL(this, objperson, "GetValue", _.ARGS.Val("MIGAgentInformation.InternalMIGPartnerID").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

                    if (_.IF(_.GT(_.NullableNUM(_.INSTR(strpersoninternalmigpartnerids, relobjmigpartnerid)), (Int16)0)))
                    {
                        flagauthorized = true;
                    }
                }
                else
                {
                    //If relObjMIGPartnerID is Null or empty, modification allowed
                    flagauthorized = true;
                }

            }

            //return
            CheckAgentHasMIGPartnerID_retVal = _.VAL(flagauthorized);
            return CheckAgentHasMIGPartnerID_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlcontext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}