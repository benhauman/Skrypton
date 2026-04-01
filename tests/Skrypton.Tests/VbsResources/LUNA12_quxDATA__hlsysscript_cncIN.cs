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

            //---------------------------------------------------------------------------------------- main ---
            _.CALLm1v0(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ProcessIn"); // call the main ü entry point
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
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "ProcessRequestMail start.");

            oMailRequest = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("mailrequest")));
            oHLServer = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("serverconnection")));

            autoReplyList = _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", "Out of Office:", "Abwesend:"));
            rfKeywords = _.VAL(_.CALLm1v1(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", "[ServiceRequest]"));
            imKeywords = _.VAL(_.CALLm1v1(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", "[Incident]"));
            cmKeywords = _.VAL(_.CALLm1v1(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", "[RFC]"));

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("mail subject:", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "subject")));

            var enumerationContent = _.ENUMERABLE(autoReplyList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject"), item, (Int16)1)), (Int16)0)))
                {
                    _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "processtext", "Out of Office AutoReply");
                    return;
                }
            }

            _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", (Int16)(-2));
            adhocMail = false;
            adhocMail = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "IsAdhocMail", _.ARGS.Ref(oMailRequest, v => { oMailRequest = v; })));

            //+++ Änderung für Workflow +++
            refNumber = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ExtractRefNumber", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject")));
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refNumber)), (Int16)0)))
            {
                caseToExtend = _.OBJ(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "GetCaseByReferenceNumber", _.ARGS.Ref(refNumber, v2 => { refNumber = v2; })));
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "RefNumber > 0");
                if (_.IF(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "IsBuiltinCase", _.ARGS.Ref(caseToExtend, v3 => { caseToExtend = v3; }))))
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "IsBuiltinCase");
                    sReportText = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "extendCaseFromMail", _.ARGS.Ref(oMailRequest, v4 => { oMailRequest = v4; }).Ref(oCaseCfg, v5 => { oCaseCfg = v5; }).Ref(oHLServer, v6 => { oHLServer = v6; }).Ref(refNumber, v7 => { refNumber = v7; })));
                    return;
                }
                else
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "NOT IsBuiltinCase");
                    if (_.IF(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "CanExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v8 => { caseToExtend = v8; }))))
                    {
                        _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "CanExtend");
                        sReportText = _.VAL(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "DoExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v9 => { caseToExtend = v9; })));
                        return;
                    }
                    else
                    {
                        _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "CanNotExtend");
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
            sReportText = _.VAL(_.CALLm1v1(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "NewWorkflowFromMail", "Request"));
            //      End If
            //		End If
            //   End If
            // End If

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "ProcessRequestMail end.");
        }
        //--------------------------------------------------------------------------------------- sub 2 ---
        public void LogText(ref object sText)
        {
            //session("worker").trace sText
            _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "processtext", _.CONCAT(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("processtext")), sText, VBScriptConstants.vbLf));
        }
        //--------------------------------------------------------------------------------------- sub 3 ---
        public void SetCaseAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "SetCaseAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "CreateScriptEngine"));

            object hlcase_vref = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("hlcase").Ref(hlcase_vref, v10 => { hlcase_vref = v10; }));
            }
            finally { hlcase = hlcase_vref; }
            object mail_vref = mail;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("mail").Ref(mail_vref, v11 => { mail_vref = v11; }));
            }
            finally { mail = mail_vref; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "ExecuteScript", _.ARGS.Ref(oScripter, v12 => { oScripter = v12; }).Ref(_env.session, v13 => { _env.session = v13; }).Val("receive"));

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

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("config")));

            oCaseCfgs = _.OBJ(_.CALLm1v1(this, oConfig ?? throw new InvalidOperationException("Reference not set:oConfig"), "GetGroup", "CaseTypes"));

            var enumerationContent2 = _.ENUMERABLE(_.CALLm1v0(this, oCaseCfgs ?? throw new InvalidOperationException("Reference not set:oCaseCfgs"), "Groups")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                oCaseType = enumerationContent2.Current;
                if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1v1(this, oCaseType ?? throw new InvalidOperationException("Reference not set:oCaseType"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "data"), _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype"))))
                {
                    oCaseCfg = _.OBJ(oCaseType);
                    _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "data")));
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
            var enumerationContent3 = _.ENUMERABLE(_.CALLm1v0(this, _.CALLm1v1(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("config")) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "GetGroup", "subject") ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "values")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                oSubjectValue = enumerationContent3.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject"), _.CALLm1v0(this, oSubjectValue ?? throw new InvalidOperationException("Reference not set:oSubjectValue"), "data"), (Int16)1)), (Int16)0)))
                {
                    _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", _.CLNG(_.CALLm1v0(this, oSubjectValue ?? throw new InvalidOperationException("Reference not set:oSubjectValue"), "Name")));
                    break;
                }
            }
            if (_.IF(_.LT(_.NullableNUM(_.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype")), (Int16)0)))
            {
                _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), "processtext", "unregistered mail subject");
                return;
            }
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("MailRequestType:", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype")));
            object oMailRequest_vref = oMailRequest;
            try
            {
                sReportText = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "createCaseFromMail", _.ARGS.Ref(oMailRequest_vref, v14 => { oMailRequest_vref = v14; }).Ref(oCaseCfg, v15 => { oCaseCfg = v15; }).Ref(oHLServer, v16 => { oHLServer = v16; })));
            }
            finally { oMailRequest = oMailRequest_vref; }
        }
        public void SetSUAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "SetSUAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "CreateScriptEngine"));

            object hlcase_vref2 = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("hlcase").Ref(hlcase_vref2, v17 => { hlcase_vref2 = v17; }));
            }
            finally { hlcase = hlcase_vref2; }
            object mail_vref2 = mail;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("mail").Ref(mail_vref2, v18 => { mail_vref2 = v18; }));
            }
            finally { mail = mail_vref2; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:session"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "ExecuteScript", _.ARGS.Ref(oScripter, v19 => { oScripter = v19; }).Ref(_env.session, v20 => { _env.session = v20; }).Val("extend"));

        }
        public void AssociateSenderToCase(ref object oMailRequest, ref object oCaseCfg, ref object oHLServer, ref object oCase)
        {
            object sMailAttributeKey = null;
            object sSearchConditionPersons = null;
            object oPersons = null;

            //
            // Suche
            //
            sMailAttributeKey = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "MailAttributeKey") ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "data"));
            sSearchConditionPersons = _.CONCAT(sMailAttributeKey, "= \"", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "SenderMail"), "\"");

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("SearchCondition = ", sSearchConditionPersons));
            oPersons = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v21 => { sSearchConditionPersons = v21; }).Val((Int16)0)));

            //
            // Baue eine Assoziation zwischen Vorgang und Anfrager
            //
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, oPersons ?? throw new InvalidOperationException("Reference not set:oPersons"), "Count")), (Int16)0)))
            {
                oPersons = VBScriptConstants.Nothing;
                // Keine Person mit der EmailAdresse gefunden !!!!
                // Besser für Auswertung mit Berichten ist ein DummyPerson
                // z.B. "email adresse unbekant" als Anfrager zu setzen
                //
                // Bitte zuerst in helpLine diese Dummy-Person anlegen !
                //
                sSearchConditionPersons = "PersonGeneral.Name = \"email adresse unbekannt\"";
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("SearchCondition2 = ", sSearchConditionPersons));
                oPersons = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v22 => { sSearchConditionPersons = v22; }).Val((Int16)0)));
                if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, oPersons ?? throw new InvalidOperationException("Reference not set:oPersons"), "Count")), (Int16)0)))
                {
                    _.CALLm1argp(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "AssociatePersons", _.ARGS.Ref(oPersons, v23 => { oPersons = v23; }));
                }
            }
            else
            {
                _.CALLm1argp(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "AssociatePersons", _.ARGS.Ref(oPersons, v24 => { oPersons = v24; }));
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

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "createCaseFromMail");

            //
            //	Erzeuge einen Vorgang
            //

            sCaseType = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "CaseType") ?? throw new InvalidOperationException("Reference not set:(_.call result)"), "data"));
            oCase = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "CreateCase", _.ARGS.Ref(sCaseType, v25 => { sCaseType = v25; })));
            oHLCase = _.OBJ(_.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "GetHLObject"));

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("case-id:", _.CSTR(_.CALLm1v0(this, oHLCase ?? throw new InvalidOperationException("Reference not set:oHLCase"), "GetID"))));

            object oMailRequest_vref2 = oMailRequest, oCaseCfg_vref = oCaseCfg, oHLServer_vref = oHLServer;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "AssociateSenderToCase", _.ARGS.Ref(oMailRequest_vref2, v26 => { oMailRequest_vref2 = v26; }).Ref(oCaseCfg_vref, v27 => { oCaseCfg_vref = v27; }).Ref(oHLServer_vref, v28 => { oHLServer_vref = v28; }).Ref(oCase, v29 => { oCase = v29; }));
            }
            finally { oMailRequest = oMailRequest_vref2; oCaseCfg = oCaseCfg_vref; oHLServer = oHLServer_vref; }

            // Setze die Attribute des Vorgangs
            //
            object oMailRequest_vref3 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "SetCaseAttributes", _.ARGS.Ref(oHLCase, v30 => { oHLCase = v30; }).Ref(oMailRequest_vref3, v31 => { oMailRequest_vref3 = v31; }));
            }
            finally { oMailRequest = oMailRequest_vref3; }

            // Gebe den Vorgang für alle User frei
            //
            _.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "Unreserve");

            // save it to the helpline server
            //
            _.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "Save");

            // Setze die Report Information
            //
            CaseRefNumber = _.VAL(_.CALLm1v5(this, oHLCase ?? throw new InvalidOperationException("Reference not set:oHLCase"), "GetValue", "CASEINFO.REFERENCENUMBER", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            sReportText = _.CONCAT(sReportText, VBScriptConstants.vbLf, "CaseType:", _.CSTR(sCaseType));
            sReportText = _.CONCAT(sReportText, VBScriptConstants.vbLf, "case-id:", _.CSTR(_.CALLm1v0(this, oHLCase ?? throw new InvalidOperationException("Reference not set:oHLCase"), "GetID")));
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

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "extendCaseFromMail");

            SearchCondition = _.CONCAT("CASEINFO.REFERENCENUMBER= ", refNumber);

            cases = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "find_Cases", _.ARGS.Ref(SearchCondition, v32 => { SearchCondition = v32; }).Val((Int16)0)));

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("cases:", _.CALLm1v0(this, cases ?? throw new InvalidOperationException("Reference not set:cases"), "count")));

            var enumerationContent4 = _.ENUMERABLE(cases).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                oCase = enumerationContent4.Current;
                object oMailRequest_vref4 = oMailRequest, oCaseCfg_vref2 = oCaseCfg, oHLServer_vref2 = oHLServer;
                try
                {
                    _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ExtendCase", _.ARGS.Ref(oCase, v33 => { oCase = v33; }).Ref(oMailRequest_vref4, v34 => { oMailRequest_vref4 = v34; }).Ref(oCaseCfg_vref2, v35 => { oCaseCfg_vref2 = v35; }).Ref(oHLServer_vref2, v36 => { oHLServer_vref2 = v36; }));
                }
                finally { oMailRequest = oMailRequest_vref4; oCaseCfg = oCaseCfg_vref2; oHLServer = oHLServer_vref2; }

                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "case extended");
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("case-id:", _.CALLm2v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "getHLObject", "getID")));
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("case-ref:", _.CSTR(refNumber)));
            }

            ExtendCaseFromMail_retVal = "";
            return ExtendCaseFromMail_retVal;
        }
        //---------------------------------------------------------------------------------------- ExtendCase ---
        public void ExtendCase(ref object oCase, ref object oMailRequest, ref object oCaseCfg, ref object oHLServer)
        {

            _.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "createSU");

            object oMailRequest_vref5 = oMailRequest, oCaseCfg_vref3 = oCaseCfg, oHLServer_vref3 = oHLServer, oCase_vref = oCase;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "AssociateSenderToCase", _.ARGS.Ref(oMailRequest_vref5, v37 => { oMailRequest_vref5 = v37; }).Ref(oCaseCfg_vref3, v38 => { oCaseCfg_vref3 = v38; }).Ref(oHLServer_vref3, v39 => { oHLServer_vref3 = v39; }).Ref(oCase_vref, v40 => { oCase_vref = v40; }));
            }
            finally { oMailRequest = oMailRequest_vref5; oCaseCfg = oCaseCfg_vref3; oHLServer = oHLServer_vref3; oCase = oCase_vref; }

            object oMailRequest_vref6 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "SetSUAttributes", _.ARGS.Val(_.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "getHLObject")).Ref(oMailRequest_vref6, v41 => { oMailRequest_vref6 = v41; }));
            }
            finally { oMailRequest = oMailRequest_vref6; }

            _.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "mergeSUs");

        }
        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object IsWFEmail(ref object subject, ref object keywordList)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "IsWFEmail called");
            var enumerationContent5 = _.ENUMERABLE(keywordList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent5.MoveNext())
                    break;
                item = enumerationContent5.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, subject, item, (Int16)1)), (Int16)0)))
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("IsWFEmail - ", item));
                    IsWFEmail_retVal = true;
                    break;
                }
                else
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("IsNotWFEmail - ", item));
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
