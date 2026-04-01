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
            _.CALLm1v0(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ProcessIn");
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
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "ProcessRequestMail start.");

            oMailRequest = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("mailrequest")));
            oHLServer = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("serverconnection")));

            autoReplyList = _.VAL(_.CALLm1argp(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", _.ARGS.ForceBrackets())); //("Out of Office:", "Abwesend:")
            rfKeywords = _.VAL(_.CALLm1argp(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", _.ARGS.ForceBrackets())); //("[ServiceRequest]", "Anfrage", "request", "Frage", "question")
            imKeywords = _.VAL(_.CALLm1argp(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", _.ARGS.ForceBrackets())); //("[Incident]", "Incident","Stoerung","Hilfe", "help")
            cmKeywords = _.VAL(_.CALLm1argp(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ARRAY", _.ARGS.ForceBrackets())); //("[RFC]", "Aenderung", "Change")

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("mail subject:", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "subject")));

            var enumerationContent = _.ENUMERABLE(autoReplyList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject"), item, (Int16)1)), (Int16)0)))
                {
                    _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "processtext", "Out of Office AutoReply");
                    return;
                }
            }

            _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", (Int16)(-2));
            adhocMail = false;
            adhocMail = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "IsAdhocMail", _.ARGS.Ref(oMailRequest, v => { oMailRequest = v; })));

            //+++ Aenderung fuer Workflow +++
            refNumber = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ExtractRefNumber", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject")));
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refNumber)), (Int16)0)))
            {
                caseToExtend = _.OBJ(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "GetCaseByReferenceNumber", _.ARGS.Ref(refNumber, v2 => { refNumber = v2; })));
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "RefNumber > 0");
                if (_.IF(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "IsBuiltinCase", _.ARGS.Ref(caseToExtend, v3 => { caseToExtend = v3; }))))
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "IsBuiltinCase");
                    sReportText = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "extendCaseFromMail", _.ARGS.Ref(oMailRequest, v4 => { oMailRequest = v4; }).Ref(oCaseCfg, v5 => { oCaseCfg = v5; }).Ref(oHLServer, v6 => { oHLServer = v6; }).Ref(refNumber, v7 => { refNumber = v7; })));
                }
                else
                {
                    _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "NOT IsBuiltinCase");
                    if (_.IF(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "CanExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v8 => { caseToExtend = v8; }))))
                    {
                        _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "CanExtend");
                        sReportText = _.VAL(_.CALLm1argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "DoExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v9 => { caseToExtend = v9; })));
                        return;
                    }
                    else
                    {
                        _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "CanNotExtend");
                        return;
                    }
                }
            }
            else
            {
                if (_.IF(_.EQ(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject")).Ref(rfKeywords, v10 => { rfKeywords = v10; })), true)))
                {
                    sReportText = _.VAL(_.CALLm1v1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "NewWorkflowFromMail", "RequestFulfillment"));
                }
                else
                {
                    if (_.IF(_.EQ(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject")).Ref(imKeywords, v11 => { imKeywords = v11; })), true)))
                    {
                        sReportText = _.VAL(_.CALLm1v1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "NewWorkflowFromMail", "IncidentManagement"));
                    }
                    else
                    {
                        if (_.IF(_.EQ(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "IsWFEmail", _.ARGS.Val(_.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject")).Ref(cmKeywords, v12 => { cmKeywords = v12; })), true)))
                        {
                            sReportText = _.VAL(_.CALLm1v1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "NewWorkflowFromMail", "ChangeManagement"));
                        }
                        else
                        {
                            if (_.IF(_.EQ(adhocMail, true)))
                            {
                                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "CreateAdhocCase", _.ARGS.Ref(oMailRequest, v13 => { oMailRequest = v13; }));
                            }
                            else
                            {
                                sReportText = _.VAL(_.CALLm1v1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "NewWorkflowFromMail", "Request"));
                            }
                        }
                    }
                }
            }

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "ProcessRequestMail end.");
        }
        //--------------------------------------------------------------------------------------- sub 2 ---
        public void LogText(ref object sText)
        {
            //session("worker").trace sText
            _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "processtext", _.CONCAT(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("processtext")), sText, VBScriptConstants.vbLf));
        }
        //--------------------------------------------------------------------------------------- sub 3 ---
        public void SetCaseAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "SetCaseAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:"), "CreateScriptEngine"));

            object hlcase_vref = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("hlcase").Ref(hlcase_vref, v14 => { hlcase_vref = v14; }));
            }
            finally { hlcase = hlcase_vref; }
            object mail_vref = mail;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("mail").Ref(mail_vref, v15 => { mail_vref = v15; }));
            }
            finally { mail = mail_vref; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:"), "ExecuteScript", _.ARGS.Ref(oScripter, v16 => { oScripter = v16; }).Ref(_env.session, v17 => { _env.session = v17; }).Val("receive"));

        }
        public object AdhocMailCfg(ref object oMailRequest)
        {
            object AdhocMailCfg_retVal = null;
            object oConfig = null;
            object oCaseCfgs = null;
            object oCaseCfg = null;
            object oCaseType = null;

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("config")));

            oCaseCfg = VBScriptConstants.Nothing;
            oCaseCfgs = _.OBJ(_.CALLm1v1(this, oConfig ?? throw new InvalidOperationException("Reference not set:oConfig"), "GetGroup", "CaseTypes"));

            var enumerationContent2 = _.ENUMERABLE(_.CALLm1v0(this, oCaseCfgs ?? throw new InvalidOperationException("Reference not set:oCaseCfgs"), "Groups")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                oCaseType = enumerationContent2.Current;
                if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1v1(this, oCaseType ?? throw new InvalidOperationException("Reference not set:oCaseType"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:"), "data"), _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype"))))
                {
                    oCaseCfg = _.OBJ(oCaseType);
                    _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:"), "data")));
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

            oConfig = _.OBJ(_.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("config")));

            oCaseCfgs = _.OBJ(_.CALLm1v1(this, oConfig ?? throw new InvalidOperationException("Reference not set:oConfig"), "GetGroup", "CaseTypes"));

            var enumerationContent3 = _.ENUMERABLE(_.CALLm1v0(this, oCaseCfgs ?? throw new InvalidOperationException("Reference not set:oCaseCfgs"), "Groups")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                oCaseType = enumerationContent3.Current;
                if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1v1(this, oCaseType ?? throw new InvalidOperationException("Reference not set:oCaseType"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:"), "data"), _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype"))))
                {
                    oCaseCfg = _.OBJ(oCaseType);
                    _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "type") ?? throw new InvalidOperationException("Reference not set:"), "data")));
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
            var enumerationContent4 = _.ENUMERABLE(_.CALLm1v0(this, _.CALLm1v1(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("config")) ?? throw new InvalidOperationException("Reference not set:"), "GetGroup", "subject") ?? throw new InvalidOperationException("Reference not set:"), "values")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent4.MoveNext())
                    break;
                oSubjectValue = enumerationContent4.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "Subject"), _.CALLm1v0(this, oSubjectValue ?? throw new InvalidOperationException("Reference not set:oSubjectValue"), "data"), (Int16)1)), (Int16)0)))
                {
                    _.SETm1a0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype", _.CLNG(_.CALLm1v0(this, oSubjectValue ?? throw new InvalidOperationException("Reference not set:oSubjectValue"), "Name")));
                    break;
                }
            }
            if (_.IF(_.LT(_.NullableNUM(_.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype")), (Int16)0)))
            {
                _.SETm0a1(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), "processtext", "unregistered mail subject");
                return;
            }
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("MailRequestType:", _.CALLm1v0(this, oMailRequest ?? throw new InvalidOperationException("Reference not set:oMailRequest"), "mailtype")));
            object oMailRequest_vref = oMailRequest;
            try
            {
                sReportText = _.VAL(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "createCaseFromMail", _.ARGS.Ref(oMailRequest_vref, v18 => { oMailRequest_vref = v18; }).Ref(oCaseCfg, v19 => { oCaseCfg = v19; }).Ref(oHLServer, v20 => { oHLServer = v20; })));
            }
            finally { oMailRequest = oMailRequest_vref; }
        }
        public void SetSUAttributes(ref object hlcase, ref object mail)
        {
            object oScripter = null;

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "SetSUAttributes");

            oScripter = _.OBJ(_.CALLm1v0(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:"), "CreateScriptEngine"));

            object hlcase_vref2 = hlcase;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("hlcase").Ref(hlcase_vref2, v21 => { hlcase_vref2 = v21; }));
            }
            finally { hlcase = hlcase_vref2; }
            object mail_vref2 = mail;
            try
            {
                _.CALLm1argp(this, oScripter ?? throw new InvalidOperationException("Reference not set:oScripter"), "AddObject", _.ARGS.Val("mail").Ref(mail_vref2, v22 => { mail_vref2 = v22; }));
            }
            finally { mail = mail_vref2; }

            _.CALLm1argp(this, _.CALLm0argp(this, _env.session ?? throw new InvalidOperationException("Reference not set:"), _.ARGS.Val("worker")) ?? throw new InvalidOperationException("Reference not set:"), "ExecuteScript", _.ARGS.Ref(oScripter, v23 => { oScripter = v23; }).Ref(_env.session, v24 => { _env.session = v24; }).Val("extend"));

        }
        public void AssociateSenderToCase(ref object oMailRequestX, ref object oCaseCfgX, ref object oHLServerX, ref object oCaseX)
        {
            object oCaseCfgZ = null;
            object sMailAttributeKey = null;
            object sSearchConditionPersons = null;
            object oPersons = null;

            object oMailRequestX_vref = oMailRequestX;
            try
            {
                oCaseCfgZ = _.OBJ(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "AdhocMailCfg", _.ARGS.Ref(oMailRequestX_vref, v25 => { oMailRequestX_vref = v25; })));
            }
            finally { oMailRequestX = oMailRequestX_vref; }
            //
            // Suche
            //
            sMailAttributeKey = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfgZ ?? throw new InvalidOperationException("Reference not set:oCaseCfgZ"), "GetValue", "MailAttributeKey") ?? throw new InvalidOperationException("Reference not set:"), "data"));
            sSearchConditionPersons = _.CONCAT(sMailAttributeKey, "= \"", _.CALLm1v0(this, oMailRequestX ?? throw new InvalidOperationException("Reference not set:oMailRequestX"), "SenderMail"), "\"");

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("SearchCondition = ", sSearchConditionPersons));
            oPersons = _.OBJ(_.CALLm1argp(this, oHLServerX ?? throw new InvalidOperationException("Reference not set:oHLServerX"), "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v26 => { sSearchConditionPersons = v26; }).Val((Int16)0)));

            //
            // Baue eine Assoziation zwischen Vorgang und Anfrager
            //
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, oPersons ?? throw new InvalidOperationException("Reference not set:oPersons"), "Count")), (Int16)0)))
            {
                oPersons = VBScriptConstants.Nothing;
                // Keine Person mit der EmailAdresse gefunden !!!!
                // Besser fuer Auswertung mit Berichten ist ein DummyPerson
                // z.B. "email adresse unbekant" als Anfrager zu setzen
                //
                // Bitte zuerst in helpLine diese Dummy-Person anlegen !
                //
                sSearchConditionPersons = "PersonGeneral.Name = \"email adresse unbekannt\"";
                _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("SearchCondition2 = ", sSearchConditionPersons));
                oPersons = _.OBJ(_.CALLm1argp(this, oHLServerX ?? throw new InvalidOperationException("Reference not set:oHLServerX"), "Find_Persons", _.ARGS.Ref(sSearchConditionPersons, v27 => { sSearchConditionPersons = v27; }).Val((Int16)0)));
                if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, oPersons ?? throw new InvalidOperationException("Reference not set:oPersons"), "Count")), (Int16)0)))
                {
                    _.CALLm1argp(this, oCaseX ?? throw new InvalidOperationException("Reference not set:oCaseX"), "AssociatePersons", _.ARGS.Ref(oPersons, v28 => { oPersons = v28; }));
                }
            }
            else
            {
                _.CALLm1argp(this, oCaseX ?? throw new InvalidOperationException("Reference not set:oCaseX"), "AssociatePersons", _.ARGS.Ref(oPersons, v29 => { oPersons = v29; }));
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
            //   Erzeuge einen Vorgang
            //

            sCaseType = _.VAL(_.CALLm1v0(this, _.CALLm1v1(this, oCaseCfg ?? throw new InvalidOperationException("Reference not set:oCaseCfg"), "GetValue", "CaseType") ?? throw new InvalidOperationException("Reference not set:"), "data"));
            oCase = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "CreateCase", _.ARGS.Ref(sCaseType, v30 => { sCaseType = v30; })));
            oHLCase = _.OBJ(_.CALLm1v0(this, oCase ?? throw new InvalidOperationException("Reference not set:oCase"), "GetHLObject"));

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("case-id:", _.CSTR(_.CALLm1v0(this, oHLCase ?? throw new InvalidOperationException("Reference not set:oHLCase"), "GetID"))));

            object oMailRequest_vref2 = oMailRequest, oCaseCfg_vref = oCaseCfg, oHLServer_vref = oHLServer;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "AssociateSenderToCase", _.ARGS.Ref(oMailRequest_vref2, v31 => { oMailRequest_vref2 = v31; }).Ref(oCaseCfg_vref, v32 => { oCaseCfg_vref = v32; }).Ref(oHLServer_vref, v33 => { oHLServer_vref = v33; }).Ref(oCase, v34 => { oCase = v34; }));
            }
            finally { oMailRequest = oMailRequest_vref2; oCaseCfg = oCaseCfg_vref; oHLServer = oHLServer_vref; }

            // Setze die Attribute des Vorgangs
            //
            object oMailRequest_vref3 = oMailRequest;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "SetCaseAttributes", _.ARGS.Ref(oHLCase, v35 => { oHLCase = v35; }).Ref(oMailRequest_vref3, v36 => { oMailRequest_vref3 = v36; }));
            }
            finally { oMailRequest = oMailRequest_vref3; }

            // Gebe den Vorgang fuer alle User frei
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

            cases = _.OBJ(_.CALLm1argp(this, oHLServer ?? throw new InvalidOperationException("Reference not set:oHLServer"), "find_Cases", _.ARGS.Ref(SearchCondition, v37 => { SearchCondition = v37; }).Val((Int16)0)));

            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", _.CONCAT("cases:", _.CALLm1v0(this, cases ?? throw new InvalidOperationException("Reference not set:cases"), "count")));

            var enumerationContent5 = _.ENUMERABLE(cases).GetEnumerator();
            while (true)
            {
                if (!enumerationContent5.MoveNext())
                    break;
                oCase = enumerationContent5.Current;
                object oMailRequest_vref4 = oMailRequest, oCaseCfg_vref2 = oCaseCfg, oHLServer_vref2 = oHLServer;
                try
                {
                    _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ExtendCase", _.ARGS.Ref(oCase, v38 => { oCase = v38; }).Ref(oMailRequest_vref4, v39 => { oMailRequest_vref4 = v39; }).Ref(oCaseCfg_vref2, v40 => { oCaseCfg_vref2 = v40; }).Ref(oHLServer_vref2, v41 => { oHLServer_vref2 = v41; }));
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
        public void ExtendCase(ref object ocaseZeC, ref object oMailRequestZeC, ref object oCaseCfg, ref object oHLServerZeC)
        {
            object oCaseCfgZeC = null; /* Undeclared in source */

            _.CALLm1v0(this, ocaseZeC ?? throw new InvalidOperationException("Reference not set:ocaseZeC"), "createSU");

            object oMailRequestZeC_vref = oMailRequestZeC, oHLServerZeC_vref = oHLServerZeC, oCaseZeC_vref = ocaseZeC;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "AssociateSenderToCase", _.ARGS.Ref(oMailRequestZeC_vref, v42 => { oMailRequestZeC_vref = v42; }).Ref(oCaseCfgZeC, v43 => { oCaseCfgZeC = v43; }).Ref(oHLServerZeC_vref, v44 => { oHLServerZeC_vref = v44; }).Ref(oCaseZeC_vref, v45 => { oCaseZeC_vref = v45; }));
            }
            finally { oMailRequestZeC = oMailRequestZeC_vref; oHLServerZeC = oHLServerZeC_vref; ocaseZeC = oCaseZeC_vref; }

            object oMailRequestZeC_vref2 = oMailRequestZeC;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "SetSUAttributes", _.ARGS.Val(_.CALLm1v0(this, ocaseZeC ?? throw new InvalidOperationException("Reference not set:ocaseZeC"), "getHLObject")).Ref(oMailRequestZeC_vref2, v46 => { oMailRequestZeC_vref2 = v46; }));
            }
            finally { oMailRequestZeC = oMailRequestZeC_vref2; }

            _.CALLm1v0(this, ocaseZeC ?? throw new InvalidOperationException("Reference not set:ocaseZeC"), "mergeSUs");

        }
        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object IsWFEmail(ref object subject, ref object keywordList)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "LogText", "IsWFEmail called");
            var enumerationContent6 = _.ENUMERABLE(keywordList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent6.MoveNext())
                    break;
                item = enumerationContent6.Current;
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
