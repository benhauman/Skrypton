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

            //---------------------------------------------------------------------------------------- main ---
            _.CALLm1v0(this, _outer, "ProcessIn"); // call the main ü entry point
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

        //--------------------------------------------------------------------------------------- sub 1 ---
        public void ProcessIn()
        {
            object oMailRequest = null;
            object oHLServer = null;
            object adhocMail = null;
            object autoReplyList = null;
            object imKeywords = null;
            object rfKeywords = null;
            object cmKeywords = null;
            object item = null;
            object sReportText = null;
            object refNumber = null;
            object caseToExtend = null;
            object oCaseCfg = null; /* Undeclared in source */
            // main ü entry point for execution
            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail start.");

            oMailRequest = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("mailrequest")));
            oHLServer = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("serverconnection")));

            autoReplyList = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.Val("Out of Office:").Val("Abwesend:")));
            rfKeywords = _.VAL(_.CALLm1v1(this, _, "ARRAY", "[ServiceRequest]"));
            imKeywords = _.VAL(_.CALLm1v1(this, _, "ARRAY", "[Incident]"));
            cmKeywords = _.VAL(_.CALLm1v1(this, _, "ARRAY", "[RFC]"));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("mail subject:", _.CALLm1v0(this, oMailRequest, "subject")));

            var enumerationContent = _.ENUMERABLE(autoReplyList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest, "Subject"), item, (Int16)1)), (Int16)0)))
                {
                    _.SET("Out of Office AutoReply", this, _env.session, null, _.ARGS.Val("processtext"));
                    return;
                }
            }

            _.SET((Int16)(-2), this, oMailRequest, "mailtype");
            adhocMail = false;
            adhocMail = _.VAL(_.CALLm1argp(this, _outer, "IsAdhocMail", _.ARGS.Ref(oMailRequest, v => { oMailRequest = v; })));

            //+++ Änderung für Workflow +++
            refNumber = _.VAL(_.CALLm1v1(this, _outer, "ExtractRefNumber", _.CALLm1v0(this, oMailRequest, "Subject")));
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refNumber)), (Int16)0)))
            {
                caseToExtend = _.OBJ(_.CALLm1argp(this, _env.session, "GetCaseByReferenceNumber", _.ARGS.Ref(refNumber, v2 => { refNumber = v2; })));
                _.CALLm1v1(this, _outer, "LogText", "RefNumber > 0");
                if (_.IF(_.CALLm1argp(this, _env.session, "IsBuiltinCase", _.ARGS.Ref(caseToExtend, v3 => { caseToExtend = v3; }))))
                {
                    _.CALLm1v1(this, _outer, "LogText", "IsBuiltinCase");
                    sReportText = _.VAL(_.CALLm1argp(this, _outer, "extendCaseFromMail", _.ARGS.Ref(oMailRequest, v4 => { oMailRequest = v4; }).Ref(oCaseCfg, v5 => { oCaseCfg = v5; }).Ref(oHLServer, v6 => { oHLServer = v6; }).Ref(refNumber, v7 => { refNumber = v7; })));
                    return;
                }
                else
                {
                    _.CALLm1v1(this, _outer, "LogText", "NOT IsBuiltinCase");
                    if (_.IF(_.CALLm1argp(this, _env.session, "CanExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v8 => { caseToExtend = v8; }))))
                    {
                        _.CALLm1v1(this, _outer, "LogText", "CanExtend");
                        sReportText = _.VAL(_.CALLm1argp(this, _env.session, "DoExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v9 => { caseToExtend = v9; })));
                        return;
                    }
                    else
                    {
                        _.CALLm1v1(this, _outer, "LogText", "CanNotExtend");
                    }
                }
            }
            //sReportText = session.NewWorkflowFromMail("AzureEvent")
            //If (IsWFEmail(oMailRequest.Subject, rfKeywords) = True) Then
            //	sReportText = session.NewWorkflowFromMail("RequestFulfillment")
            //Else
            //  If (IsWFEmail(oMailRequest.Subject, imKeywords) = True) Then
            // 		sReportText = session.NewWorkflowFromMail("IncidentManagement")
            //	Else
            //   	If (IsWFEmail(oMailRequest.Subject, cmKeywords) = True) Then
            //		sReportText = session.NewWorkflowFromMail("ChangeManagement")
            //	Else
            //      If adhocMail = True Then
            //      	CreateAdhocCase oMailRequest
            //      Else
            sReportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "Request"));
            //      End If
            //		End If
            //   End If
            // End If

            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail end.");
        }

        //--------------------------------------------------------------------------------------- sub 2 ---
        public void LogText(ref object sText)
        {
            //session("worker").trace sText
            _.SET(_.CONCAT(_.CALLm0argp(this, _env.session, _.ARGS.Val("processtext")), sText, VBScriptConstants.vbLf), this, _env.session, null, _.ARGS.Val("processtext"));
        }

        //--------------------------------------------------------------------------------------- sub 3 ---
        public void SetCaseAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer, "LogText", "SetCaseAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "CreateScriptEngine"));

            object byrefalias = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("hlcase").Ref(byrefalias, v10 => { byrefalias = v10; }));
            }
            finally { hlcase = byrefalias; }
            object byrefalias2 = mail;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("mail").Ref(byrefalias2, v11 => { byrefalias2 = v11; }));
            }
            finally { mail = byrefalias2; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "ExecuteScript", _.ARGS.Ref(oScripter, v12 => { oScripter = v12; }).Ref(_env.session, v13 => { _env.session = v13; }).Val("receive"));

        }

        public object IsAdhocMail(ref object oMailRequest)
        {
            object IsAdhocMail_retVal = null;
            object bRegisteredMailType = null;
            object oConfig = null;
            object oCaseCfgs = null;
            object oCaseCfg = null;
            object oCaseType = null;
            //
            //	Suche die Konfiguration für diesen Vorgangstypen
            //
            bRegisteredMailType = false;

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("config")));

            oCaseCfgs = _.OBJ(_.CALLm1v1(this, oConfig, "GetGroup", "CaseTypes"));

            var enumerationContent2 = _.ENUMERABLE(_.CALLm1v0(this, oCaseCfgs, "Groups")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                oCaseType = enumerationContent2.Current;
                if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1v1(this, oCaseType, "GetValue", "type"), "data"), _.CALLm1v0(this, oMailRequest, "mailtype"))))
                {
                    oCaseCfg = _.OBJ(oCaseType);
                    _.SET(_.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "type"), "data")), this, oMailRequest, "mailtype");
                    bRegisteredMailType = true;
                    break;
                }
            }

            IsAdhocMail_retVal = _.VAL(bRegisteredMailType);
            return IsAdhocMail_retVal;
        }

        public void CreateAdhocCase(ref object oMailRequest)
        {
            object oSubjectValue = null;
            object sReportText = null; /* Undeclared in source */
            object oCaseCfg = null; /* Undeclared in source */
            object oHLServer = null; /* Undeclared in source */
            //
            // Suche die Objektdefinition anhand des Betreffs in der E-Mail
            //
            var enumerationContent3 = _.ENUMERABLE(_.CALLm1v0(this, _.CALLm1v1(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("config")), "GetGroup", "subject"), "values")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                oSubjectValue = enumerationContent3.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest, "Subject"), _.CALLm1v0(this, oSubjectValue, "data"), (Int16)1)), (Int16)0)))
                {
                    _.SET(_.CLNG(_.CALLm1v0(this, oSubjectValue, "Name")), this, oMailRequest, "mailtype");
                    break;
                }
            }
            if (_.IF(_.LT(_.NullableNUM(_.CALLm1v0(this, oMailRequest, "mailtype")), (Int16)0)))
            {
                _.SET("unregistered mail subject", this, _env.session, null, _.ARGS.Val("processtext"));
                return;
            }
            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("MailRequestType:", _.CALLm1v0(this, oMailRequest, "mailtype")));
            object byrefalias3 = oMailRequest;
            try
            {
                sReportText = _.VAL(_.CALLm1argp(this, _outer, "createCaseFromMail", _.ARGS.Ref(byrefalias3, v14 => { byrefalias3 = v14; }).Ref(oCaseCfg, v15 => { oCaseCfg = v15; }).Ref(oHLServer, v16 => { oHLServer = v16; })));
            }
            finally { oMailRequest = byrefalias3; }
        }

        public void SetSUAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer, "LogText", "SetSUAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "CreateScriptEngine"));

            object byrefalias4 = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("hlcase").Ref(byrefalias4, v17 => { byrefalias4 = v17; }));
            }
            finally { hlcase = byrefalias4; }
            object byrefalias5 = mail;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("mail").Ref(byrefalias5, v18 => { byrefalias5 = v18; }));
            }
            finally { mail = byrefalias5; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "ExecuteScript", _.ARGS.Ref(oScripter, v19 => { oScripter = v19; }).Ref(_env.session, v20 => { _env.session = v20; }).Val("extend"));

        }

        public void AssociateSenderToCase(ref object oMailRequest, ref object oCaseCfg, ref object oHLServer, ref object oCase)
        {
            object sMailAttributeKey = null;
            object sSearchConditionPersons = null;
            object oPersons = null;

            //
            // Suche
            //
            sMailAttributeKey = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "MailAttributeKey"), "data"));
            sSearchConditionPersons = _.CONCAT(sMailAttributeKey, "= \"", _.CALLm1v0(this, oMailRequest, "SenderMail"), "\"");

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("SearchCondition = ", sSearchConditionPersons));
            oPersons = _.OBJ(_.CALLm1argp(this, oHLServer, "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v21 => { sSearchConditionPersons = v21; }).Val((Int16)0)));

            //
            // Baue eine Assoziation zwischen Vorgang und Anfrager
            //
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, oPersons, "Count")), (Int16)0)))
            {
                oPersons = VBScriptConstants.Nothing;
                // Keine Person mit der EmailAdresse gefunden !!!!
                // Besser für Auswertung mit Berichten ist ein DummyPerson
                // z.B. "email adresse unbekant" als Anfrager zu setzen
                //
                // Bitte zuerst in helpLine diese Dummy-Person anlegen !
                //
                sSearchConditionPersons = "PersonGeneral.Name = \"email adresse unbekannt\"";
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("SearchCondition2 = ", sSearchConditionPersons));
                oPersons = _.OBJ(_.CALLm1argp(this, oHLServer, "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v22 => { sSearchConditionPersons = v22; }).Val((Int16)0)));
                if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, oPersons, "Count")), (Int16)0)))
                {
                    _.CALLm1argp(this, oCase, "AssociatePersons", _.ARGS.Ref(oPersons, v23 => { oPersons = v23; }));
                }
            }
            else
            {
                _.CALLm1argp(this, oCase, "AssociatePersons", _.ARGS.Ref(oPersons, v24 => { oPersons = v24; }));
            }

        }

        //---------------------------------------------------------------------------------------- createCaseFromMail ---
        public object CreateCaseFromMail(ref object oMailRequest, ref object oCaseCfg, ref object oHLServer)
        {
            object CreateCaseFromMail_retVal = null;
            object sCaseType = null;
            object oCase = null;
            object oHLCase = null;
            object CaseRefNumber = null;
            object sReportText = null;

            _.CALLm1v1(this, _outer, "LogText", "createCaseFromMail");

            //
            //	Erzeuge einen Vorgang
            //

            sCaseType = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "CaseType"), "data"));
            oCase = _.OBJ(_.CALLm1argp(this, oHLServer, "CreateCase", _.ARGS.Ref(sCaseType, v25 => { sCaseType = v25; })));
            oHLCase = _.OBJ(_.CALLm1v0(this, oCase, "GetHLObject"));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-id:", _.CSTR(_.CALLm1v0(this, oHLCase, "GetID"))));

            object byrefalias6 = oMailRequest, byrefalias7 = oCaseCfg, byrefalias8 = oHLServer;
            try
            {
                _.CALLm1argp(this, _outer, "AssociateSenderToCase", _.ARGS.Ref(byrefalias6, v26 => { byrefalias6 = v26; }).Ref(byrefalias7, v27 => { byrefalias7 = v27; }).Ref(byrefalias8, v28 => { byrefalias8 = v28; }).Ref(oCase, v29 => { oCase = v29; }));
            }
            finally { oMailRequest = byrefalias6; oCaseCfg = byrefalias7; oHLServer = byrefalias8; }

            // Setze die Attribute des Vorgangs
            //
            object byrefalias9 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer, "SetCaseAttributes", _.ARGS.Ref(oHLCase, v30 => { oHLCase = v30; }).Ref(byrefalias9, v31 => { byrefalias9 = v31; }));
            }
            finally { oMailRequest = byrefalias9; }

            // Gebe den Vorgang für alle User frei
            //
            _.CALLm1v0(this, oCase, "Unreserve");

            // save it to the helpline server
            //
            _.CALLm1v0(this, oCase, "Save");

            // Setze die Report Information
            //
            CaseRefNumber = _.VAL(_.CALLm1argp(this, oHLCase, "GetValue", _.ARGS.Val("CASEINFO.REFERENCENUMBER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            sReportText = _.CONCAT(sReportText, VBScriptConstants.vbLf, "CaseType:", _.CSTR(sCaseType));
            sReportText = _.CONCAT(sReportText, VBScriptConstants.vbLf, "case-id:", _.CSTR(_.CALLm1v0(this, oHLCase, "GetID")));
            sReportText = _.CONCAT(sReportText, VBScriptConstants.vbLf, "case-ref:", _.CSTR(CaseRefNumber));

            CreateCaseFromMail_retVal = _.VAL(sReportText);
            return CreateCaseFromMail_retVal;
        }

        //---------------------------------------------------------------------------------------- extractRefNumber ---
        public object ExtractRefNumber(ref object subject)
        {
            object ExtractRefNumber_retVal = null;
            object refNum = null;
            object startPos = null;
            object endPos = null;

            refNum = "";

            startPos = _.VAL(_.INSTR((Int16)1, subject, "[#", (Int16)1));
            if (_.IF(_.GT(_.NullableNUM(startPos), (Int16)0)))
            {
                startPos = _.ADD(startPos, (Int16)2); // skip "[#"

                endPos = _.VAL(_.INSTR(startPos, subject, "]", (Int16)1));

                if (_.IF(_.GT(_.NullableNUM(endPos), (Int16)0)))
                {
                    refNum = _.VAL(_.MID(subject, startPos, _.SUBT(endPos, startPos)));
                }
            }

            ExtractRefNumber_retVal = _.VAL(refNum);
            return ExtractRefNumber_retVal;
        }

        //---------------------------------------------------------------------------------------- extendCaseFromMail ---
        public object ExtendCaseFromMail(ref object oMailRequest, ref object oCaseCfg, ref object oHLServer, ref object refNumber)
        {
            object ExtendCaseFromMail_retVal = null;
            object SearchCondition = null;
            object cases = null;
            object oCase = null;

            _.CALLm1v1(this, _outer, "LogText", "extendCaseFromMail");

            SearchCondition = _.CONCAT("CASEINFO.REFERENCENUMBER= ", refNumber);

            cases = _.OBJ(_.CALLm1argp(this, oHLServer, "find_Cases", _.ARGS.Ref(SearchCondition, v32 => { SearchCondition = v32; }).Val((Int16)0)));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("cases:", _.CALLm1v0(this, cases, "count")));

            var enumerationContent4 = _.ENUMERABLE(cases).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                oCase = enumerationContent4.Current;
                object byrefalias10 = oMailRequest, byrefalias11 = oCaseCfg, byrefalias12 = oHLServer;
                try
                {
                    _.CALLm1argp(this, _outer, "ExtendCase", _.ARGS.Ref(oCase, v33 => { oCase = v33; }).Ref(byrefalias10, v34 => { byrefalias10 = v34; }).Ref(byrefalias11, v35 => { byrefalias11 = v35; }).Ref(byrefalias12, v36 => { byrefalias12 = v36; }));
                }
                finally { oMailRequest = byrefalias10; oCaseCfg = byrefalias11; oHLServer = byrefalias12; }

                _.CALLm1v1(this, _outer, "LogText", "case extended");
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-id:", _.CALLm2v0(this, oCase, "getHLObject", "getID")));
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-ref:", _.CSTR(refNumber)));
            }

            ExtendCaseFromMail_retVal = "";
            return ExtendCaseFromMail_retVal;
        }

        //---------------------------------------------------------------------------------------- ExtendCase ---
        public void ExtendCase(ref object oCase, ref object oMailRequest, ref object oCaseCfg, ref object oHLServer)
        {

            _.CALLm1v0(this, oCase, "createSU");

            object byrefalias13 = oMailRequest, byrefalias14 = oCaseCfg, byrefalias15 = oHLServer, byrefalias16 = oCase;
            try
            {
                _.CALLm1argp(this, _outer, "AssociateSenderToCase", _.ARGS.Ref(byrefalias13, v37 => { byrefalias13 = v37; }).Ref(byrefalias14, v38 => { byrefalias14 = v38; }).Ref(byrefalias15, v39 => { byrefalias15 = v39; }).Ref(byrefalias16, v40 => { byrefalias16 = v40; }));
            }
            finally { oMailRequest = byrefalias13; oCaseCfg = byrefalias14; oHLServer = byrefalias15; oCase = byrefalias16; }

            object byrefalias17 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer, "SetSUAttributes", _.ARGS.Val(_.CALLm1v0(this, oCase, "getHLObject")).Ref(byrefalias17, v41 => { byrefalias17 = v41; }));
            }
            finally { oMailRequest = byrefalias17; }

            _.CALLm1v0(this, oCase, "mergeSUs");

        }

        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object IsWFEmail(ref object subject, ref object keywordList)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            _.CALLm1v1(this, _outer, "LogText", "IsWFEmail called");
            var enumerationContent5 = _.ENUMERABLE(keywordList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent5.MoveNext())
                    break;
                item = enumerationContent5.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, subject, item, (Int16)1)), (Int16)0)))
                {
                    _.CALLm1v1(this, _outer, "LogText", _.CONCAT("IsWFEmail - ", item));
                    IsWFEmail_retVal = true;
                    break;
                }
                else
                {
                    _.CALLm1v1(this, _outer, "LogText", _.CONCAT("IsNotWFEmail - ", item));
                    IsWFEmail_retVal = false;
                }
            }
            return IsWFEmail_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object session { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
