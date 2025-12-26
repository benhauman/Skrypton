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
        public GlobalReferences Go(EnvironmentReferences env)
        {
            var _env = env ?? throw new ArgumentNullException(nameof(env));
            var _outer = new GlobalReferences(_, _env);

            //---------------------------------------------------------------------------------------- main ---
            _.CALL(this, _outer, "ProcessIn");
            return _outer;
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
            object retVal1 = null;
            object autoreplylist = null;
            object item = null;
            object retval = null;
            retval = false;
            autoreplylist = _.VAL(_.CALL(this, _, "ARRAY", _.ARGS.Val("Out of Office:").Val("Abwesend:")));

            var enumerationContent2 = _.ENUMERABLE(autoreplylist).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                item = enumerationContent2.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailsubject, item, (Int16)1)), (Int16)0)))
                {
                    retval = true;
                }
            }
            retVal1 = _.VAL(retval);
            return retVal1;
        }

        //--------------------------------------------------------------------------------------- TryExtendCase ---
        public object tryextendcase(ref object mailsubject)
        {
            object retVal3 = null;
            object refnumber = null;
            object casetoextend = null;
            object reporttext = null;
            object retval = null;
            retval = false;

            object byrefalias4 = mailsubject;
            try
            {
                refnumber = _.VAL(_.CALL(this, _outer, "ExtractRefNumber", _.ARGS.Ref(byrefalias4, v5 => { byrefalias4 = v5; })));
            }
            finally { mailsubject = byrefalias4; }
            if (_.IF(_.GT(_.NullableNUM(_.LEN(refnumber)), (Int16)0)))
            {
                _.CALL(this, _outer, "LogText", _.ARGS.Val("RefNumber > 0"));
                casetoextend = _.OBJ(_.CALL(this, _env.session, "GetCaseByReferenceNumber", _.ARGS.Ref(refnumber, v6 => { refnumber = v6; })));
                if (_.IF(_.CALL(this, _env.session, "CanExtendWorkflowCase", _.ARGS.Ref(casetoextend, v7 => { casetoextend = v7; }))))
                {
                    _.CALL(this, _outer, "LogText", _.ARGS.Val("CanExtend"));
                    reporttext = _.VAL(_.CALL(this, _env.session, "DoExtendWorkflowCase", _.ARGS.Ref(casetoextend, v8 => { casetoextend = v8; })));
                    _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v9 => { reporttext = v9; }));
                    retval = true;
                }
            }
            retVal3 = _.VAL(retval);
            return retVal3;
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

            bool ifResult13;
            object byrefalias12 = mailsubject;
            try
            {
                ifResult13 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias12, v14 => { byrefalias12 = v14; }).Ref(rfkeywords, v15 => { rfkeywords = v15; })), true));
            }
            finally { mailsubject = byrefalias12; }
            if (ifResult13)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("RequestFulfillment")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v16 => { reporttext = v16; }));
                return;
            }
            bool ifResult20;
            object byrefalias19 = mailsubject;
            try
            {
                ifResult20 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias19, v21 => { byrefalias19 = v21; }).Ref(imkeywords, v22 => { imkeywords = v22; })), true));
            }
            finally { mailsubject = byrefalias19; }
            if (ifResult20)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("IncidentManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v23 => { reporttext = v23; }));
                return;
            }
            bool ifResult27;
            object byrefalias26 = mailsubject;
            try
            {
                ifResult27 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias26, v28 => { byrefalias26 = v28; }).Ref(cmkeywords, v29 => { cmkeywords = v29; })), true));
            }
            finally { mailsubject = byrefalias26; }
            if (ifResult27)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("ChangeManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v30 => { reporttext = v30; }));
                return;
            }
            bool ifResult34;
            object byrefalias33 = mailsubject;
            try
            {
                ifResult34 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias33, v35 => { byrefalias33 = v35; }).Ref(fmkeywords, v36 => { fmkeywords = v36; })), true));
            }
            finally { mailsubject = byrefalias33; }
            if (ifResult34)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("FacilityIncidentManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v37 => { reporttext = v37; }));
                return;
            }
            bool ifResult41;
            object byrefalias40 = mailsubject;
            try
            {
                ifResult41 = _.IF(_.EQ(_.CALL(this, _outer, "IsWFEmail", _.ARGS.Ref(byrefalias40, v42 => { byrefalias40 = v42; }).Ref(hrkeywords, v43 => { hrkeywords = v43; })), true));
            }
            finally { mailsubject = byrefalias40; }
            if (ifResult41)
            {
                reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("HRRequestManagement")));
                _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v44 => { reporttext = v44; }));
                return;
            }
            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("Request")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v45 => { reporttext = v45; }));
        }

        //--------------------------------------------------------------------------------------- StartNewFMWorkflow ---
        public void startnewfmworkflow(ref object mailsubject)
        {
            object reporttext = null;

            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("FacilityIncidentManagement")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v46 => { reporttext = v46; }));
        }

        //--------------------------------------------------------------------------------------- StartNewHRWorkflow ---
        public void startnewhrworkflow(ref object mailsubject)
        {
            object reporttext = null;

            reporttext = _.VAL(_.CALL(this, _env.session, "NewWorkflowFromMail", _.ARGS.Val("HRRequestManagement")));
            _.CALL(this, _outer, "LogText", _.ARGS.Ref(reporttext, v47 => { reporttext = v47; }));
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
            object retVal48 = null;
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
            retVal48 = _.VAL(refnum);
            return retVal48;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object isfmmail(ref object mailto)
        {
            object retVal49 = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsFMMail called"));
            retval = false;
            if (_.IF(_.EQ(_.NullableSTR(mailto), "haustechnik@helplinedemo.de")))
            {
                retval = true;
            }

            retVal49 = _.VAL(retval);
            return retVal49;
        }

        //--------------------------------------------------------------------------------------- IsFMMail ---
        public object ishrmail(ref object mailto)
        {
            object retVal50 = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsHRMail called"));
            retval = false;
            if (_.IF(_.EQ(_.NullableSTR(mailto), "personal@helplinedemo.de")))
            {
                retval = true;
            }

            retVal50 = _.VAL(retval);
            return retVal50;
        }

        //---------------------------------------------------------------------------------------- IsWorkflowEmail ---
        public object iswfemail(ref object mailsubject, ref object keywordlist)
        {
            object retVal51 = null;
            object item = null;
            object retval = null;
            _.CALL(this, _outer, "LogText", _.ARGS.Val("IsWFEmail called"));
            retval = false;

            var enumerationContent52 = _.ENUMERABLE(keywordlist).GetEnumerator();
            while (true)
            {
                if (!enumerationContent52.MoveNext())
                    break;
                item = enumerationContent52.Current;
                if (_.IF(_.GT(_.NullableNUM(_.INSTR((Int16)1, mailsubject, item, (Int16)1)), (Int16)0)))
                {
                    _.CALL(this, _outer, "LogText", _.ARGS.Val(_.CONCAT("IsWFEmail - ", item)));
                    retval = true;
                    break;
                }
            }
            retVal51 = _.VAL(retval);
            return retVal51;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object session { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}