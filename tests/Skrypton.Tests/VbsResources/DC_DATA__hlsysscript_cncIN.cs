using System;
using System.Collections;
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
            _.CALLm1v0(this, _outer, "ProcessIn");
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
            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail start.");

            oMailRequest = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("mailrequest")));
            oHLServer = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("serverconnection")));

            autoReplyList = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.ForceBrackets())); //("Out of Office:", "Abwesend:")
            rfKeywords = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.ForceBrackets())); //("[ServiceRequest]", "Anfrage", "request", "Frage", "question")
            imKeywords = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.ForceBrackets())); //("[Incident]", "Incident","Stoerung","Hilfe", "help")
            cmKeywords = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.ForceBrackets())); //("[RFC]", "Aenderung", "Change")

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("mail subject:", _.CALLm1v0(this, oMailRequest, "subject")));

            var enumerationContent = _.ENUMERABLE(autoReplyList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest, "Subject"), item, (Int16)1)), (Int16)0)))
                {
                    _.SETm0a1(this, _env.session, "processtext", "Out of Office AutoReply");
                    return;
                }
            }

            _.SETm1a0(this, oMailRequest, "mailtype", (Int16)(-2));
            adhocMail = false;
            adhocMail = _.VAL(_.CALLm1argp(this, _outer, "IsAdhocMail", _.ARGS.Ref(oMailRequest, v => { oMailRequest = v; })));

            //+++ Aenderung fuer Workflow +++
            refNumber = _.VAL(_.CALLm1v1(this, _outer, "ExtractRefNumber", _.CALLm1v0(this, oMailRequest, "Subject")));
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refNumber)), (Int16)0)))
            {
                caseToExtend = _.OBJ(_.CALLm1argp(this, _env.session, "GetCaseByReferenceNumber", _.ARGS.Ref(refNumber, v2 => { refNumber = v2; })));
                _.CALLm1v1(this, _outer, "LogText", "RefNumber > 0");
                if (_.IF(_.CALLm1argp(this, _env.session, "IsBuiltinCase", _.ARGS.Ref(caseToExtend, v3 => { caseToExtend = v3; }))))
                {
                    _.CALLm1v1(this, _outer, "LogText", "IsBuiltinCase");
                    sReportText = _.VAL(_.CALLm1argp(this, _outer, "extendCaseFromMail", _.ARGS.Ref(oMailRequest, v4 => { oMailRequest = v4; }).Ref(oCaseCfg, v5 => { oCaseCfg = v5; }).Ref(oHLServer, v6 => { oHLServer = v6; }).Ref(refNumber, v7 => { refNumber = v7; })));
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
                        return;
                    }
                }
            }
            else
            {
                if (_.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest, "Subject")).Ref(rfKeywords, v10 => { rfKeywords = v10; })), true)))
                {
                    sReportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "RequestFulfillment"));
                }
                else
                {
                    if (_.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest, "Subject")).Ref(imKeywords, v11 => { imKeywords = v11; })), true)))
                    {
                        sReportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "IncidentManagement"));
                    }
                    else
                    {
                        if (_.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest, "Subject")).Ref(cmKeywords, v12 => { cmKeywords = v12; })), true)))
                        {
                            sReportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "ChangeManagement"));
                        }
                        else
                        {
                            if (_.IF(_.EQ(adhocMail, true)))
                            {
                                _.CALLm1argp(this, _outer, "CreateAdhocCase", _.ARGS.Ref(oMailRequest, v13 => { oMailRequest = v13; }));
                            }
                            else
                            {
                                sReportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "Request"));
                            }
                        }
                    }
                }
            }

            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail end.");
        }

        //--------------------------------------------------------------------------------------- sub 2 ---
        public void LogText(ref object sText)
        {
            //session("worker").trace sText
            _.SETm0a1(this, _env.session, "processtext", _.CONCAT(_.CALLm0argp(this, _env.session, _.ARGS.Val("processtext")), sText, VBScriptConstants.vbLf));
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
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("hlcase").Ref(byrefalias, v14 => { byrefalias = v14; }));
            }
            finally { hlcase = byrefalias; }
            object byrefalias2 = mail;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("mail").Ref(byrefalias2, v15 => { byrefalias2 = v15; }));
            }
            finally { mail = byrefalias2; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "ExecuteScript", _.ARGS.Ref(oScripter, v16 => { oScripter = v16; }).Ref(_env.session, v17 => { _env.session = v17; }).Val("receive"));

        }

        public object AdhocMailCfg(ref object oMailRequest)
        {
            object AdhocMailCfg_retVal = null;
            object oConfig = null;
            object oCaseCfgs = null;
            object oCaseCfg = null;
            object oCaseType = null;

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("config")));

            oCaseCfg = VBScriptConstants.Nothing;
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
                    _.SETm1a0(this, oMailRequest, "mailtype", _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "type"), "data")));
                    break;
                }
            }

            AdhocMailCfg_retVal = _.OBJ(oCaseCfg);
            return AdhocMailCfg_retVal;
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
            //      Suche die Konfiguration fuer diesen Vorgangstypen
            //
            bRegisteredMailType = false;

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("config")));

            oCaseCfgs = _.OBJ(_.CALLm1v1(this, oConfig, "GetGroup", "CaseTypes"));

            var enumerationContent3 = _.ENUMERABLE(_.CALLm1v0(this, oCaseCfgs, "Groups")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                oCaseType = enumerationContent3.Current;
                if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1v1(this, oCaseType, "GetValue", "type"), "data"), _.CALLm1v0(this, oMailRequest, "mailtype"))))
                {
                    oCaseCfg = _.OBJ(oCaseType);
                    _.SETm1a0(this, oMailRequest, "mailtype", _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "type"), "data")));
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
            var enumerationContent4 = _.ENUMERABLE(_.CALLm1v0(this, _.CALLm1v1(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("config")), "GetGroup", "subject"), "values")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                oSubjectValue = enumerationContent4.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest, "Subject"), _.CALLm1v0(this, oSubjectValue, "data"), (Int16)1)), (Int16)0)))
                {
                    _.SETm1a0(this, oMailRequest, "mailtype", _.CLNG(_.CALLm1v0(this, oSubjectValue, "Name")));
                    break;
                }
            }
            if (_.IF(_.LT(_.NullableNUM(_.CALLm1v0(this, oMailRequest, "mailtype")), (Int16)0)))
            {
                _.SETm0a1(this, _env.session, "processtext", "unregistered mail subject");
                return;
            }
            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("MailRequestType:", _.CALLm1v0(this, oMailRequest, "mailtype")));
            object byrefalias3 = oMailRequest;
            try
            {
                sReportText = _.VAL(_.CALLm1argp(this, _outer, "createCaseFromMail", _.ARGS.Ref(byrefalias3, v18 => { byrefalias3 = v18; }).Ref(oCaseCfg, v19 => { oCaseCfg = v19; }).Ref(oHLServer, v20 => { oHLServer = v20; })));
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
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("hlcase").Ref(byrefalias4, v21 => { byrefalias4 = v21; }));
            }
            finally { hlcase = byrefalias4; }
            object byrefalias5 = mail;
            try
            {
                _.CALLm1argp(this, oScripter, "AddObject", _.ARGS.Val("mail").Ref(byrefalias5, v22 => { byrefalias5 = v22; }));
            }
            finally { mail = byrefalias5; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session, _.ARGS.Val("worker")), "ExecuteScript", _.ARGS.Ref(oScripter, v23 => { oScripter = v23; }).Ref(_env.session, v24 => { _env.session = v24; }).Val("extend"));

        }

        public void AssociateSenderToCase(ref object oMailRequestX, ref object oCaseCfgX, ref object oHLServerX, ref object oCaseX)
        {
            object oCaseCfgZ = null;
            object sMailAttributeKey = null;
            object sSearchConditionPersons = null;
            object oPersons = null;

            object byrefalias6 = oMailRequestX;
            try
            {
                oCaseCfgZ = _.OBJ(_.CALLm1argp(this, _outer, "AdhocMailCfg", _.ARGS.Ref(byrefalias6, v25 => { byrefalias6 = v25; })));
            }
            finally { oMailRequestX = byrefalias6; }
            //
            // Suche
            //
            sMailAttributeKey = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfgZ, "GetValue", "MailAttributeKey"), "data"));
            sSearchConditionPersons = _.CONCAT(sMailAttributeKey, "= \"", _.CALLm1v0(this, oMailRequestX, "SenderMail"), "\"");

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("SearchCondition = ", sSearchConditionPersons));
            oPersons = _.OBJ(_.CALLm1argp(this, oHLServerX, "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v26 => { sSearchConditionPersons = v26; }).Val((Int16)0)));

            //
            // Baue eine Assoziation zwischen Vorgang und Anfrager
            //
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, oPersons, "Count")), (Int16)0)))
            {
                oPersons = VBScriptConstants.Nothing;
                // Keine Person mit der EmailAdresse gefunden !!!!
                // Besser fuer Auswertung mit Berichten ist ein DummyPerson
                // z.B. "email adresse unbekant" als Anfrager zu setzen
                //
                // Bitte zuerst in helpLine diese Dummy-Person anlegen !
                //
                sSearchConditionPersons = "PersonGeneral.Name = \"email adresse unbekannt\"";
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("SearchCondition2 = ", sSearchConditionPersons));
                oPersons = _.OBJ(_.CALLm1argp(this, oHLServerX, "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v27 => { sSearchConditionPersons = v27; }).Val((Int16)0)));
                if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, oPersons, "Count")), (Int16)0)))
                {
                    _.CALLm1argp(this, oCaseX, "AssociatePersons", _.ARGS.Ref(oPersons, v28 => { oPersons = v28; }));
                }
            }
            else
            {
                _.CALLm1argp(this, oCaseX, "AssociatePersons", _.ARGS.Ref(oPersons, v29 => { oPersons = v29; }));
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
            //   Erzeuge einen Vorgang
            //

            sCaseType = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg, "GetValue", "CaseType"), "data"));
            oCase = _.OBJ(_.CALLm1argp(this, oHLServer, "CreateCase", _.ARGS.Ref(sCaseType, v30 => { sCaseType = v30; })));
            oHLCase = _.OBJ(_.CALLm1v0(this, oCase, "GetHLObject"));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-id:", _.CSTR(_.CALLm1v0(this, oHLCase, "GetID"))));

            object byrefalias7 = oMailRequest, byrefalias8 = oCaseCfg, byrefalias9 = oHLServer;
            try
            {
                _.CALLm1argp(this, _outer, "AssociateSenderToCase", _.ARGS.Ref(byrefalias7, v31 => { byrefalias7 = v31; }).Ref(byrefalias8, v32 => { byrefalias8 = v32; }).Ref(byrefalias9, v33 => { byrefalias9 = v33; }).Ref(oCase, v34 => { oCase = v34; }));
            }
            finally { oMailRequest = byrefalias7; oCaseCfg = byrefalias8; oHLServer = byrefalias9; }

            // Setze die Attribute des Vorgangs
            //
            object byrefalias10 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer, "SetCaseAttributes", _.ARGS.Ref(oHLCase, v35 => { oHLCase = v35; }).Ref(byrefalias10, v36 => { byrefalias10 = v36; }));
            }
            finally { oMailRequest = byrefalias10; }

            // Gebe den Vorgang fuer alle User frei
            //
            _.CALLm1v0(this, oCase, "Unreserve");

            // save it to the helpline server
            //
            _.CALLm1v0(this, oCase, "Save");

            // Setze die Report Information
            //
            CaseRefNumber = _.VAL(_.CALLm1v5(this, oHLCase, "GetValue", "CASEINFO.REFERENCENUMBER", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

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

            cases = _.OBJ(_.CALLm1argp(this, oHLServer, "find_Cases", _.ARGS.Ref(SearchCondition, v37 => { SearchCondition = v37; }).Val((Int16)0)));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("cases:", _.CALLm1v0(this, cases, "count")));

            var enumerationContent5 = _.ENUMERABLE(cases).GetEnumerator();
            while (true)
            {
                if (!enumerationContent5.MoveNext())
                    break;
                oCase = enumerationContent5.Current;
                object byrefalias11 = oMailRequest, byrefalias12 = oCaseCfg, byrefalias13 = oHLServer;
                try
                {
                    _.CALLm1argp(this, _outer, "ExtendCase", _.ARGS.Ref(oCase, v38 => { oCase = v38; }).Ref(byrefalias11, v39 => { byrefalias11 = v39; }).Ref(byrefalias12, v40 => { byrefalias12 = v40; }).Ref(byrefalias13, v41 => { byrefalias13 = v41; }));
                }
                finally { oMailRequest = byrefalias11; oCaseCfg = byrefalias12; oHLServer = byrefalias13; }

                _.CALLm1v1(this, _outer, "LogText", "case extended");
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-id:", _.CALLm2v0(this, oCase, "getHLObject", "getID")));
                _.CALLm1v1(this, _outer, "LogText", _.CONCAT("case-ref:", _.CSTR(refNumber)));
            }

            ExtendCaseFromMail_retVal = "";
            return ExtendCaseFromMail_retVal;
        }

        //---------------------------------------------------------------------------------------- ExtendCase ---
        public void ExtendCase(ref object ocaseZeC, ref object oMailRequestZeC, ref object oCaseCfg, ref object oHLServerZeC)
        {
            object oCaseCfgZeC = null; /* Undeclared in source */

            _.CALLm1v0(this, ocaseZeC, "createSU");

            object byrefalias14 = oMailRequestZeC, byrefalias15 = oHLServerZeC, byrefalias16 = ocaseZeC;
            try
            {
                _.CALLm1argp(this, _outer, "AssociateSenderToCase", _.ARGS.Ref(byrefalias14, v42 => { byrefalias14 = v42; }).Ref(oCaseCfgZeC, v43 => { oCaseCfgZeC = v43; }).Ref(byrefalias15, v44 => { byrefalias15 = v44; }).Ref(byrefalias16, v45 => { byrefalias16 = v45; }));
            }
            finally { oMailRequestZeC = byrefalias14; oHLServerZeC = byrefalias15; ocaseZeC = byrefalias16; }

            object byrefalias17 = oMailRequestZeC;
            try
            {
                _.CALLm1argp(this, _outer, "SetSUAttributes", _.ARGS.Val(_.CALLm1v0(this, ocaseZeC, "getHLObject")).Ref(byrefalias17, v46 => { byrefalias17 = v46; }));
            }
            finally { oMailRequestZeC = byrefalias17; }

            _.CALLm1v0(this, ocaseZeC, "mergeSUs");

        }

        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object IsWFEmail(ref object subject, ref object keywordList)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            _.CALLm1v1(this, _outer, "LogText", "IsWFEmail called");
            var enumerationContent6 = _.ENUMERABLE(keywordList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent6.MoveNext())
                    break;
                item = enumerationContent6.Current;
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
