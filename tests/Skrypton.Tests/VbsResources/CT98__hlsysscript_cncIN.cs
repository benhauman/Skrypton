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

        //--------------------------------------------------------------------------------------- ProcessIn ---
        public void ProcessIn()
        {
            object mailRequest = null;
            object extendCaseSuccess = null;
            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail start.");

            mailRequest = _.OBJ(_.CALLm0argp(this, _env.session, _.ARGS.Val("mailrequest")));

            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("mail subject: ", _.CALLm1v0(this, mailRequest, "subject")));
            _.CALLm1v1(this, _outer, "LogText", _.CONCAT("mail To: ", _.CALLm1v0(this, mailRequest, "To")));

            if (_.IF(_.CALLm1v1(this, _outer, "IsAutoReplyMail", _.CALLm1v0(this, mailRequest, "Subject"))))
            {
                _.CALLm1v1(this, _outer, "LogText", "Out of Office AutoReply");
                return;
            }

            extendCaseSuccess = _.VAL(_.CALLm1v1(this, _outer, "TryExtendCase", _.CALLm1v0(this, mailRequest, "Subject")));
            if (_.IF(_.EQ(extendCaseSuccess, false)))
            {
                _.CALLm1v1(this, _outer, "LogText", "Extend case failed. Start new process");
                if (_.IF(_.CALLm1v1(this, _outer, "IsFMMail", _.CALLm1v0(this, mailRequest, "To"))))
                {
                    _.CALLm1v1(this, _outer, "StartNewFMWorkflow", _.CALLm1v0(this, mailRequest, "Subject"));
                }
                else if (_.IF(_.CALLm1v1(this, _outer, "IsHRMail", _.CALLm1v0(this, mailRequest, "To"))))
                {
                    _.CALLm1v1(this, _outer, "StartNewHRWorkflow", _.CALLm1v0(this, mailRequest, "Subject"));
                }
                else
                {
                    _.CALLm1v1(this, _outer, "StartNewWorkflow", _.CALLm1v0(this, mailRequest, "Subject"));
                }
            }

            _.CALLm1v1(this, _outer, "LogText", "ProcessRequestMail end.");
        }

        //--------------------------------------------------------------------------------------- IsAutoReplyMail ---
        public object IsAutoReplyMail(ref object mailSubject)
        {
            object IsAutoReplyMail_retVal = null;
            object autoReplyList = null;
            object item = null;
            object retVal = null;
            retVal = false;
            autoReplyList = _.VAL(_.CALLm1v2(this, _, "ARRAY", "Out of Office:", "Abwesend:"));

            var enumerationContent = _.ENUMERABLE(autoReplyList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailSubject, item, (Int16)1)), (Int16)0)))
                {
                    retVal = true;
                }
            }
            IsAutoReplyMail_retVal = _.VAL(retVal);
            return IsAutoReplyMail_retVal;
        }

        //--------------------------------------------------------------------------------------- TryExtendCase ---
        public object TryExtendCase(ref object mailSubject)
        {
            object TryExtendCase_retVal = null;
            object refNumber = null;
            object caseToExtend = null;
            object reportText = null;
            object retVal = null;
            retVal = false;

            object byrefalias = mailSubject;
            try
            {
                refNumber = _.VAL(_.CALLm1argp(this, _outer, "ExtractRefNumber", _.ARGS.Ref(byrefalias, v => { byrefalias = v; })));
            }
            finally { mailSubject = byrefalias; }
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refNumber)), (Int16)0)))
            {
                _.CALLm1v1(this, _outer, "LogText", "RefNumber > 0");
                caseToExtend = _.OBJ(_.CALLm1argp(this, _env.session, "GetCaseByReferenceNumber", _.ARGS.Ref(refNumber, v2 => { refNumber = v2; })));
                if (_.IF(_.CALLm1argp(this, _env.session, "CanExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v3 => { caseToExtend = v3; }))))
                {
                    _.CALLm1v1(this, _outer, "LogText", "CanExtend");
                    reportText = _.VAL(_.CALLm1argp(this, _env.session, "DoExtendWorkflowCase", _.ARGS.Ref(caseToExtend, v4 => { caseToExtend = v4; })));
                    _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v5 => { reportText = v5; }));
                    retVal = true;
                }
            }
            TryExtendCase_retVal = _.VAL(retVal);
            return TryExtendCase_retVal;
        }

        //--------------------------------------------------------------------------------------- StartNewWorkflow ---
        public void StartNewWorkflow(ref object mailSubject)
        {
            object imKeywords = null;
            object rfKeywords = null;
            object cmKeywords = null;
            object fmKeywords = null;
            object hrKeywords = null;
            object reportText = null;
            rfKeywords = _.VAL(_.CALLm1v5(this, _, "ARRAY", "[ServiceRequest]", "Anfrage", "request", "Frage", "question"));
            imKeywords = _.VAL(_.CALLm1v5(this, _, "ARRAY", "[Incident]", "Incident", "Störung", "Hilfe", "help"));
            cmKeywords = _.VAL(_.CALLm1v3(this, _, "ARRAY", "[RFC]", "Änderung", "Change"));
            fmKeywords = _.VAL(_.CALLm1v3(this, _, "ARRAY", "[Facility]", "Haustechnik", "FM"));
            hrKeywords = _.VAL(_.CALLm1v2(this, _, "ARRAY", "[HR]", "Personal"));

            bool ifResult;
            object byrefalias2 = mailSubject;
            try
            {
                ifResult = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias2, v8 => { byrefalias2 = v8; }).Ref(rfKeywords, v9 => { rfKeywords = v9; })), true));
            }
            finally { mailSubject = byrefalias2; }
            if (ifResult)
            {
                reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "RequestFulfillment"));
                _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v10 => { reportText = v10; }));
                return;
            }
            bool ifResult2;
            object byrefalias3 = mailSubject;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias3, v13 => { byrefalias3 = v13; }).Ref(imKeywords, v14 => { imKeywords = v14; })), true));
            }
            finally { mailSubject = byrefalias3; }
            if (ifResult2)
            {
                reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "IncidentManagement"));
                _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v15 => { reportText = v15; }));
                return;
            }
            bool ifResult3;
            object byrefalias4 = mailSubject;
            try
            {
                ifResult3 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias4, v18 => { byrefalias4 = v18; }).Ref(cmKeywords, v19 => { cmKeywords = v19; })), true));
            }
            finally { mailSubject = byrefalias4; }
            if (ifResult3)
            {
                reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "ChangeManagement"));
                _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v20 => { reportText = v20; }));
                return;
            }
            bool ifResult4;
            object byrefalias5 = mailSubject;
            try
            {
                ifResult4 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias5, v23 => { byrefalias5 = v23; }).Ref(fmKeywords, v24 => { fmKeywords = v24; })), true));
            }
            finally { mailSubject = byrefalias5; }
            if (ifResult4)
            {
                reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "FacilityIncidentManagement"));
                _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v25 => { reportText = v25; }));
                return;
            }
            bool ifResult5;
            object byrefalias6 = mailSubject;
            try
            {
                ifResult5 = _.IF(_.EQ(_.CALLm1argp(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias6, v28 => { byrefalias6 = v28; }).Ref(hrKeywords, v29 => { hrKeywords = v29; })), true));
            }
            finally { mailSubject = byrefalias6; }
            if (ifResult5)
            {
                reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "HRRequestManagement"));
                _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v30 => { reportText = v30; }));
                return;
            }
            reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "Request"));
            _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v31 => { reportText = v31; }));
        }

        //--------------------------------------------------------------------------------------- StartNewFMWorkflow ---
        public void StartNewFMWorkflow(ref object mailSubject)
        {
            object reportText = null;

            reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "FacilityIncidentManagement"));
            _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v32 => { reportText = v32; }));
        }

        //--------------------------------------------------------------------------------------- StartNewHRWorkflow ---
        public void StartNewHRWorkflow(ref object mailSubject)
        {
            object reportText = null;

            reportText = _.VAL(_.CALLm1v1(this, _env.session, "NewWorkflowFromMail", "HRRequestManagement"));
            _.CALLm1argp(this, _outer, "LogText", _.ARGS.Ref(reportText, v33 => { reportText = v33; }));
        }

        //--------------------------------------------------------------------------------------- LogText ---
        public void LogText(ref object sText)
        {
            //Uncomment to enable logging
            _.SET(_.CONCAT(_.CALLm0argp(this, _env.session, _.ARGS.Val("processtext")), sText, VBScriptConstants.vbNewLine), this, _env.session, null, _.ARGS.Val("processtext"));
        }

        //---------------------------------------------------------------------------------------- ExtractRefNumber ---
        public object ExtractRefNumber(ref object mailSubject)
        {
            object ExtractRefNumber_retVal = null;
            object refNum = null;
            object startPos = null;
            object endPos = null;
            refNum = "";

            startPos = _.VAL(_.INSTR((Int16)1, mailSubject, "[#", (Int16)1));
            if (_.IF(_.GT(_.NullableNUM(startPos), (Int16)0)))
            {
                startPos = _.ADD(startPos, (Int16)2); // skip "[#"
                endPos = _.VAL(_.INSTR(startPos, mailSubject, "]", (Int16)1));
                if (_.IF(_.GT(_.NullableNUM(endPos), (Int16)0)))
                {
                    refNum = _.VAL(_.MID(mailSubject, startPos, _.SUBT(endPos, startPos)));
                }
            }
            ExtractRefNumber_retVal = _.VAL(refNum);
            return ExtractRefNumber_retVal;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object IsFMMail(ref object mailTo)
        {
            object IsFMMail_retVal = null;
            object retVal = null;
            _.CALLm1v1(this, _outer, "LogText", "IsFMMail called");
            retVal = false;
            if (_.IF(_.EQ(_.NullableSTR(mailTo), "haustechnik@helplinedemo.de")))
            {
                retVal = true;
            }

            IsFMMail_retVal = _.VAL(retVal);
            return IsFMMail_retVal;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object IsHRMail(ref object mailTo)
        {
            object IsHRMail_retVal = null;
            object retVal = null;
            _.CALLm1v1(this, _outer, "LogText", "IsHRMail called");
            retVal = false;
            if (_.IF(_.EQ(_.NullableSTR(mailTo), "personal@helplinedemo.de")))
            {
                retVal = true;
            }

            IsHRMail_retVal = _.VAL(retVal);
            return IsHRMail_retVal;
        }

        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object IsWFEmail(ref object mailSubject, ref object keywordList)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            object retVal = null;
            _.CALLm1v1(this, _outer, "LogText", "IsWFEmail called");
            retVal = false;

            var enumerationContent2 = _.ENUMERABLE(keywordList).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                item = enumerationContent2.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailSubject, item, (Int16)1)), (Int16)0)))
                {
                    _.CALLm1v1(this, _outer, "LogText", _.CONCAT("IsWFEmail - ", item));
                    retVal = true;
                    break;
                }
            }
            IsWFEmail_retVal = _.VAL(retVal);
            return IsWFEmail_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object session { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
