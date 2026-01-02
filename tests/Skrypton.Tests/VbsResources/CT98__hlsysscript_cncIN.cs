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
            _.CALL(this, _outer, "ProcessIn");
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
        public void processin()
        {
            object mailrequest = null;
            object extendcasesuccess = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("ProcessRequestMail start."));

            mailrequest = _.OBJ(_.CALL(this, _env.session, _.ARGS.Val("mailrequest")));

            _.CALL(this, _outer, "LogText", _.ARGS.Val(_.CONCAT("mail subject: ", _.CALL(this, mailrequest, "subject"))));
            _.CALL(this, _outer, "LogText", _.ARGS.Val(_.CONCAT("mail To: ", _.CALL(this, mailrequest, "To"))));

            if (_.IF(_.CALL(this, _outer, "IsAutoReplyMail", _.ARGS.Val(_.CALL(this, mailrequest, "Subject")))))
            {
                _.CALL(this, _outer, "LogText", _.ARGS.Val("Out of Office AutoReply"));
                return;
            }

            extendcasesuccess = _.VAL(_.CALL(this, _outer, "TryExtendCase", _.ARGS.Val(_.CALL(this, mailrequest, "Subject"))));
            if (_.IF(_.EQ(extendcasesuccess, false)))
            {
                _.CALL(this, _outer, "LogText", _.ARGS.Val("Extend case failed. Start new process"));
                if (_.IF(_.CALL(this, _outer, "IsFMMail", _.ARGS.Val(_.CALL(this, mailrequest, "To")))))
                {
                    _.CALL(this, _outer, "StartNewFMWorkflow", _.ARGS.Val(_.CALL(this, mailrequest, "Subject")));
                }
                else if (_.IF(_.CALL(this, _outer, "IsHRMail", _.ARGS.Val(_.CALL(this, mailrequest, "To")))))
                {
                    _.CALL(this, _outer, "StartNewHRWorkflow", _.ARGS.Val(_.CALL(this, mailrequest, "Subject")));
                }
                else
                {
                    _.CALL(this, _outer, "StartNewWorkflow", _.ARGS.Val(_.CALL(this, mailrequest, "Subject")));
                }
            }

            _.CALL(this, _outer, "LogText", _.ARGS.Val("ProcessRequestMail end."));
        }

        //--------------------------------------------------------------------------------------- IsAutoReplyMail ---
        public object isautoreplymail(ref object mailsubject)
        {
            object IsAutoReplyMail_retVal = null;
            object autoreplylist = null;
            object item = null;
            object retval = null;
            retval = false;
            autoreplylist = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("Out of Office:").Val("Abwesend:")));

            var enumerationContent = _.ENUMERABLE(autoreplylist).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailsubject, item, (Int16)1)), (Int16)0)))
                {
                    retval = true;
                }
            }
            IsAutoReplyMail_retVal = _.VAL(retval);
            return IsAutoReplyMail_retVal;
        }

        //--------------------------------------------------------------------------------------- TryExtendCase ---
        public object tryextendcase(ref object mailsubject)
        {
            object TryExtendCase_retVal = null;
            object refnumber = null;
            object casetoextend = null;
            object reporttext = null;
            object retval = null;
            retval = false;

            object byrefalias = mailsubject;
            try
            {
                refnumber = _.VAL(_.CALL(this, _outer, "ExtractRefNumber", _.ARGS.Ref(byrefalias, v => { byrefalias = v; })));
            }
            finally { mailsubject = byrefalias; }
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refnumber)), (Int16)0)))
            {
                _.CALL(this, _outer, "LogText", _.ARGS.Val("RefNumber > 0"));
                casetoextend = _.OBJ(_.CALL(this, _env.session, "GetCaseByReferenceNumber", _.ARGS.Ref(refnumber, v2 => { refnumber = v2; })));
                if (_.IF(_.CALL(this, _env.session, "CanExtendWorkflowCase", _.ARGS.Ref(casetoextend, v3 => { casetoextend = v3; }))))
                {
                    _.CALL(this, _outer, "LogText", _.ARGS.Val("CanExtend"));
                    reporttext = _.VAL(_.CALL(this, _env.session, "DoExtendWorkflowCase", _.ARGS.Ref(casetoextend, v4 => { casetoextend = v4; })));
                    _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v5 => { reporttext = v5; }));
                    retval = true;
                }
            }
            TryExtendCase_retVal = _.VAL(retval);
            return TryExtendCase_retVal;
        }

        //--------------------------------------------------------------------------------------- StartNewWorkflow ---
        public void startnewworkflow(ref object mailsubject)
        {
            object imkeywords = null;
            object rfkeywords = null;
            object cmkeywords = null;
            object fmkeywords = null;
            object hrkeywords = null;
            object reporttext = null;
            rfkeywords = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("[ServiceRequest]").Val("Anfrage").Val("request").Val("Frage").Val("question")));
            imkeywords = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("[Incident]").Val("Incident").Val("Störung").Val("Hilfe").Val("help")));
            cmkeywords = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("[RFC]").Val("Änderung").Val("Change")));
            fmkeywords = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("[Facility]").Val("Haustechnik").Val("FM")));
            hrkeywords = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("[HR]").Val("Personal")));

            bool ifResult;
            object byrefalias2 = mailsubject;
            try
            {
                ifResult = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias2, v8 => { byrefalias2 = v8; }).Ref(rfkeywords, v9 => { rfkeywords = v9; })), true));
            }
            finally { mailsubject = byrefalias2; }
            if (ifResult)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("RequestFulfillment")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v10 => { reporttext = v10; }));
                return;
            }
            bool ifResult2;
            object byrefalias3 = mailsubject;
            try
            {
                ifResult2 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias3, v13 => { byrefalias3 = v13; }).Ref(imkeywords, v14 => { imkeywords = v14; })), true));
            }
            finally { mailsubject = byrefalias3; }
            if (ifResult2)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("IncidentManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v15 => { reporttext = v15; }));
                return;
            }
            bool ifResult3;
            object byrefalias4 = mailsubject;
            try
            {
                ifResult3 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias4, v18 => { byrefalias4 = v18; }).Ref(cmkeywords, v19 => { cmkeywords = v19; })), true));
            }
            finally { mailsubject = byrefalias4; }
            if (ifResult3)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("ChangeManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v20 => { reporttext = v20; }));
                return;
            }
            bool ifResult4;
            object byrefalias5 = mailsubject;
            try
            {
                ifResult4 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias5, v23 => { byrefalias5 = v23; }).Ref(fmkeywords, v24 => { fmkeywords = v24; })), true));
            }
            finally { mailsubject = byrefalias5; }
            if (ifResult4)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("FacilityIncidentManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v25 => { reporttext = v25; }));
                return;
            }
            bool ifResult5;
            object byrefalias6 = mailsubject;
            try
            {
                ifResult5 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias6, v28 => { byrefalias6 = v28; }).Ref(hrkeywords, v29 => { hrkeywords = v29; })), true));
            }
            finally { mailsubject = byrefalias6; }
            if (ifResult5)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("HRRequestManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v30 => { reporttext = v30; }));
                return;
            }
            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("Request")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v31 => { reporttext = v31; }));
        }

        //--------------------------------------------------------------------------------------- StartNewFMWorkflow ---
        public void startnewfmworkflow(ref object mailsubject)
        {
            object reporttext = null;

            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("FacilityIncidentManagement")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v32 => { reporttext = v32; }));
        }

        //--------------------------------------------------------------------------------------- StartNewHRWorkflow ---
        public void startnewhrworkflow(ref object mailsubject)
        {
            object reporttext = null;

            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("HRRequestManagement")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v33 => { reporttext = v33; }));
        }

        //--------------------------------------------------------------------------------------- LogText ---
        public void logtext(ref object stext)
        {
            //Uncomment to enable logging
            _.SET(_.CONCAT(_.CALL(this, _env.session, _.ARGS.Val("processtext")), stext, VBScriptConstants.vbNewLine), this, _env.session, null, _.ARGS.Val("processtext"));
        }

        //---------------------------------------------------------------------------------------- ExtractRefNumber ---
        public object extractrefnumber(ref object mailsubject)
        {
            object ExtractRefNumber_retVal = null;
            object refnum = null;
            object startpos = null;
            object endpos = null;
            refnum = "";

            startpos = _.VAL(_.INSTR((Int16)1, mailsubject, "[#", (Int16)1));
            if (_.IF(_.GT(_.NullableNUM(startpos), (Int16)0)))
            {
                startpos = _.ADD(startpos, (Int16)2); // skip "[#"
                endpos = _.VAL(_.INSTR(startpos, mailsubject, "]", (Int16)1));
                if (_.IF(_.GT(_.NullableNUM(endpos), (Int16)0)))
                {
                    refnum = _.VAL(_.MID(mailsubject, startpos, _.SUBT(endpos, startpos)));
                }
            }
            ExtractRefNumber_retVal = _.VAL(refnum);
            return ExtractRefNumber_retVal;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object isfmmail(ref object mailto)
        {
            object IsFMMail_retVal = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsFMMail called"));
            retval = false;
            if (_.IF(_.EQ(_.NullableSTR(mailto), "haustechnik@helplinedemo.de")))
            {
                retval = true;
            }

            IsFMMail_retVal = _.VAL(retval);
            return IsFMMail_retVal;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object ishrmail(ref object mailto)
        {
            object IsHRMail_retVal = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsHRMail called"));
            retval = false;
            if (_.IF(_.EQ(_.NullableSTR(mailto), "personal@helplinedemo.de")))
            {
                retval = true;
            }

            IsHRMail_retVal = _.VAL(retval);
            return IsHRMail_retVal;
        }

        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object iswfemail(ref object mailsubject, ref object keywordlist)
        {
            object IsWFEmail_retVal = null;
            object item = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsWFEmail called"));
            retval = false;

            var enumerationContent2 = _.ENUMERABLE(keywordlist).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                item = enumerationContent2.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailsubject, item, (Int16)1)), (Int16)0)))
                {
                    _.CALL(this, _outer, "LogText", _.ARGS.Val(_.CONCAT("IsWFEmail - ", item)));
                    retval = true;
                    break;
                }
            }
            IsWFEmail_retVal = _.VAL(retval);
            return IsWFEmail_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object session { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}