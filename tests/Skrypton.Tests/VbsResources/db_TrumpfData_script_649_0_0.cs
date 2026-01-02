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

            _outer.hlasc_softwarelicensegroupview = 110944;
            _outer.hlasc_softwarelicensefolderview = 110941;
            //----------------------------------------------------------------------------------------------------------
            //Globale Konstanten fuer freie Assoziationsdefinitionen

            //XML-Export Neuanlage

            //XML-Export Incident wegen Eleminierung

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
        }

        internal object hlasc_softwarelicensefolderview { get; set; }
        internal object hlasc_softwarelicensegroupview { get; set; }

        //----------------------------------------------------------------------------------------------------------
        //Globale Funktion zur Initialisierung der Datei hlStartITIL2.dll.
        //Diese DLL-Datei beinhaltet alle globalen Funktionen und Prozeduren,
        //die innerhalb der Start ITIL Konfiguration verwendet werden.
        //Diese Funktion darf nicht aus dem gloabeln Script entfernt werden !
        //Global Function for initializing the file hlStartITIL2.dll.
        //This dll file contains any global functions and subs used for the
        //Start ITIL configuration.
        //Do not remove this function from the global script !
        //Copyright (C) 1994-2006 PMCS GmbH & Co.
        public object hlitil2()
        {
            object hlITIL2_retVal = null;
            hlITIL2_retVal = _.OBJ(_.CREATEOBJECT("hlStartITIL2.Global"));
            _.CALL(this, _outer, "hlITIL2", "SelfCheck", _.ARGS.Ref(_env.hlcontext, v => { _env.hlcontext = v; }));
            return hlITIL2_retVal;
        }

        //----------------------------------------------------------------------------------------------------------
        //Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
        public void trace(ref object hlcontext, ref object text)
        {
            object byrefalias = text;
            try
            {
                _.CALL(this, hlcontext, "trace", _.ARGS.Val((Int16)1).Ref(byrefalias, v2 => { byrefalias = v2; }));
            }
            finally { text = byrefalias; }
        }

        //----------------------------------------------------------------------------------------------------------
        //Funktion InfoMail
        //Zum Aufrufen aus EBL-Skripten von Vorgaengen
        public void infomail(ref object hlcontext, ref object hlcase, ref object subject, ref object mailsender, ref object receiver, ref object cc, ref object body, ref object sendattachments)
        {
            object email = null;
            object attachids = null;
            object attachid = null;
            object attachment = null;
            object mailattachment = null;

            email = _.OBJ(_.CALL(this, hlcontext, "CreateMail"));

            //Falls der Parameter <SendAttachmnets> beim Aufruf "1" ist, werden Anhaenge mitversandt
            if (_.IF(_.EQ(_.CBOOL(sendattachments), true)))
            {
                attachment = VBScriptConstants.Nothing;
                attachids = _.VAL(_.CALL(this, hlcase, "GetAttachmentKeys", _.ARGS.Val("HLOBJECTINFO.ATTACHMENT").Val((Int16)0)));
                var enumerationContent = _.ENUMERABLE(attachids).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent.MoveNext())
                        break;
                    attachid = enumerationContent.Current;
                    attachment = _.OBJ(_.CALL(this, hlcase, "GetAttachment", _.ARGS.Val("HLOBJECTINFO.ATTACHMENT").Ref(attachid, v3 => { attachid = v3; }).Val((Int16)0)));
                    if (_.IF(_.GT(_.NullableNUM(_.CALL(this, attachment, "Size")), (Int16)0)))
                    {
                        mailattachment = VBScriptConstants.Nothing;
                        mailattachment = _.OBJ(_.CALL(this, email, "AddAttachment"));
                        _.SET(_.VAL(_.CALL(this, attachment, "name")), this, mailattachment, "name");
                        _.SET(_.VAL(_.CALL(this, attachment, "data")), this, mailattachment, "data");
                    }
                }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(mailsender), "")))
            {
                _.SET(_.VAL(mailsender), this, email, "SenderMail");
            }
            _.SET(_.VAL(receiver), this, email, "To");
            _.SET(_.VAL(subject), this, email, "Subject");
            _.SET(_.VAL(body), this, email, "Body");
            if (_.IF(_.NOTEQ(_.NullableSTR(cc), "")))
            {
                _.SET(_.VAL(cc), this, email, "CC");
            }
            _.CALL(this, hlcontext, "SendRequestMail", _.ARGS.Ref(email, v4 => { email = v4; }));
        }

        //----------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------
        public void createsubject(ref object hlcontext, ref object survey, ref object hlcaller)
        {
            object language = null;
            language = _.VAL(_.CALL(this, hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.Language").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableSTR(language), "LanguageGerman")))
            {
                _.CALL(this, survey, "SetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("Umfrage zur Service-Leistung ihres Support-Teams"));
            }
            else
            {
                _.CALL(this, survey, "SetValue", _.ARGS.Val("CaseGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("Survey about the Service-Quality from your Support-Team"));
            }
        }

        //----------------------------------------------------------------------------------------------------------
        public void invitesurveyemail(ref object hlcontext, ref object hlcase, ref object hlcaller)
        {
            object suidx = null;
            object mailrequest = null;
            object strcrlf = null;
            object creationdate = null;
            object datum = null;
            object subject = null;
            object body = null;
            object refnumber = null;
            object portallink = null;
            object surname = null;
            object letteraddress = null;
            object anrede = null;
            object personaddress = null;
            object language = null;
            object email = null;
            object emailadress = null;
            //Email an den Anfrager eines Survey-Vorgangs, um diesen zur Teilnahme an der
            //Umfrage aufzufordern.
            //Email to Requester of a Survey-Case to invite him to take part on the survey
            object byrefalias2 = hlcase, byrefalias3 = hlcontext;
            try
            {
                suidx = _.VAL(_.CALL(this, _outer, "hlITIL2", "GetLastSUIdx", _.ARGS.Ref(byrefalias2, v5 => { byrefalias2 = v5; }).Ref(byrefalias3, v6 => { byrefalias3 = v6; })));
            }
            finally { hlcase = byrefalias2; hlcontext = byrefalias3; }
            mailrequest = _.VAL(_.CALL(this, hlcase, "GetValue", _.ARGS.Val("CaseGeneral.DefaultNotification").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.AND(_.EQ(_.NullableSTR(mailrequest), "DefaultNotificationEmail"), _.EQ(_.NullableNUM(suidx), (Int16)1))))
            {
                strcrlf = _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10));
                refnumber = _.VAL(_.CALL(this, hlcase, "GetValue", _.ARGS.Val("CASEINFO.REFERENCENUMBER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                portallink = "http://localhost/helplineportal/";
                surname = _.VAL(_.CALL(this, hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.PersonSurname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                letteraddress = _.VAL(_.CALL(this, hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.ShortLetterAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                anrede = "Sehr geehrte Damen und Herren,";
                personaddress = "Dear Mrs./Ms. or Mr.,";
                language = _.VAL(_.CALL(this, hlcaller, "GetValue", _.ARGS.Val("PersonGeneral.Language").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

                if (_.IF(_.EQ(_.NullableSTR(language), "LanguageGerman")))
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Herr/Frau";
                    }
                    anrede = _.CONCAT("Sehr geehrte(r) ", _.CSTR(letteraddress), " ", _.CSTR(surname), ",");

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    creationdate = _.VAL(_.CALL(this, hlcase, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    datum = _.VAL(_.MID(creationdate, (Int16)1, (Int16)10));
                    subject = "Umfrage zur Service-Leistung ihres Support-Teams";

                    //Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
                    //If the case was recorded, the requester receives the according information
                    body = _.CSTR(anrede);
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Wir wollen besser werden!");
                    body = _.CONCAT(body, strcrlf, "Dazu benoetigen wir Ihre Unterstuetzung und Ihr Feedback.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Sie wurden am ", datum, " durch ein Zufallsverfahren ausgewaehlt, an einer Umfrage zu den Service-Leistungen Ihres Support-Teams teilzunehmen.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Die Teilnahme ist freiwillig und erfolgt ueber das helpLine Portal.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Rufen Sie im Browser bitte folgende URL auf:");
                    body = _.CONCAT(body, strcrlf, portallink, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Klicken Sie unter 'Ihre Anfragen' auf den Eintrag 'Umfragen'. ");
                    body = _.CONCAT(body, "Dort finden Sie das Umfrage-Formular mit der Nummer ", refnumber, ". ");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Wir freuen uns sehr, wenn Sie sich die Zeit nehmen, die Fragen zu beantworten.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Wir bedanken uns fuer Ihre Unterstuetzung!");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, strcrlf, "Mit freundlichen Gruessen");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Ihr Support Team");
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Mrs./Ms./Mr.";
                    }
                    personaddress = _.CONCAT("Dear ", _.ADD(_.CSTR(letteraddress), " "), _.CSTR(surname), ",");

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    creationdate = _.VAL(_.CALL(this, hlcase, "GetValue", _.ARGS.Val("HLOBJECTINFO.CREATIONTIME").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    datum = _.VAL(_.MID(creationdate, (Int16)1, (Int16)10));
                    subject = "Survey about the Service-Quality from your Support-Team";

                    //Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
                    //If the case was recorded, the requester receives the according information
                    body = _.CSTR(personaddress);
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "We would like to improve the efficiency of Service-Support!");
                    body = _.CONCAT(body, strcrlf, "Therefore we need your assistance and your feedback.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "You where chosen by random on ", datum, " to take part on the Survey.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "The participation on the survey is voluntarily. You can take part on the survey via the helpLine Portal.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Start your Browser and choose the following URL:");
                    body = _.CONCAT(body, strcrlf, portallink, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Then klick 'Survey' in the menue 'Your Requests'. ", strcrlf);
                    body = _.CONCAT(body, "There, you will find the Questionnaire with the reference number ", refnumber, ". ");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "It would be nice, if you invest your time to response the questions.");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "We thank you for your assistance!");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, strcrlf, "With best regards");
                    body = _.CONCAT(body, strcrlf, strcrlf);
                    body = _.CONCAT(body, "Yours Support Team");
                }

                email = _.OBJ(_.CALL(this, hlcontext, "CreateMail"));

                //Ermittle die Emailadresse des Anfragers
                //Detect email adress of requester
                emailadress = _.VAL(_.CALL(this, hlcaller, "GetValue", _.ARGS.Val("PersonInformation.EmailAddress").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                if (_.IF(_.EQ(_.NullableSTR(emailadress), "")))
                {
                    emailadress = "Username@yourcompany.com";
                    subject = "Diese EMail konnte nicht zugestellt werden";
                    body = "Die Mail fuer die Anfragenummer ";
                    body = _.CONCAT(body, _.CALL(this, hlcase, "GetValue", _.ARGS.Val("CASEINFO.REFERENCENUMBER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                    body = _.CONCAT(body, " konnte wegen einer fehlenden E-Mail Adresse nicht zugestellt werden.");
                }
                _.SET(_.VAL(emailadress), this, email, "To");
                _.SET(_.VAL(subject), this, email, "Subject");
                _.SET(_.VAL(body), this, email, "Body");
                _.CALL(this, hlcontext, "SendRequestMail", _.ARGS.Ref(email, v7 => { email = v7; }));
            }
        }

        //----------------------------------------------------------------------------------------------------------
        //Diese Funktion steuert den SystemTask wenn dieser im Vorgangstyp Task konfiguriert wurde.
        //This function controls a SystemTask if it had been configured within the casetype Task.
        public void mytask1(ref object hlcontext)
        {
            object hlobj = null;
            object lcid = null;
            object langid = null;
            object priority = null;
            object tasktype = null;
            object subject = null;
            object description = null;
            object exoperation = null;
            object assignedgroup = null;
            object assignedperson = null;
            object team = null;
            object newtask = null;
            object hascontent = null;
            object contentids = null;
            object contentid = null;
            object newcontentid = null;
            object assignedgroupwf = null;
            object assignedpersonwf = null;
            object descriptionwf = null;
            object subjectwf = null;
            object assets = null;
            object asset = null;
            object refnumber = null;
            hlobj = _.OBJ(_.CALL(this, hlcontext, "GetCurrentObject", _.ARGS.ForceBrackets()));
            lcid = (Int16)0;
            lcid = _.VAL(_.CALL(this, hlcontext, "GetLocaleID"));
            langid = (Int16)0;
            langid = _.VAL(_.CALL(this, hlcontext, "LangIDFromLCID", _.ARGS.Ref(lcid, v8 => { lcid = v8; })));

            //Gesetzte Daten aus dem aktuellen Task auslesen, diese werden dem zu erzeugenden Systemtask mitgegeben.
            //Read setted data of current task and take them into the created Systemtask.
            priority = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CaseClassificationAttribute.Priority").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            tasktype = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskGeneral.TaskType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            subject = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            description = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            exoperation = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            assignedgroup = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CaseSpecialRouting.AssignedGroup").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            assignedperson = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CaseSpecialRouting.AssignedPerson").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            team = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            newtask = _.OBJ(_.CALL(this, hlcontext, "createobject", _.ARGS.Val("Task")));

            _.CALL(this, newtask, "SetValue", _.ARGS.Val("CaseClassificationAttribute.Priority").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(priority, v9 => { priority = v9; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskGeneral.TaskType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(tasktype, v10 => { tasktype = v10; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(subject, v11 => { subject = v11; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(team, v12 => { team = v12; }));

            hascontent = _.VAL(_.CALL(this, hlobj, "HasContent", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA").Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.NOTEQ(_.NullableNUM(hascontent), (Int16)0)))
            {
                subjectwf = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskDesignWorkflow.FlagWorkflowSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
                _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskDesignWorkflow.FlagWorkflowSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(subjectwf, v13 => { subjectwf = v13; }));
                contentids = _.VAL(_.CALL(this, hlobj, "GetContentIDs", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA").Val((Int16)0)));
                _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskWorkflowAttribute.WorkflowStep").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1));
                var enumerationContent2 = _.ENUMERABLE(contentids).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent2.MoveNext())
                        break;
                    contentid = enumerationContent2.Current;
                    assignedgroupwf = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup").Val((Int16)0).Ref(contentid, v14 => { contentid = v14; }).Val((Int16)0).Val((Int16)0)));
                    assignedpersonwf = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson").Val((Int16)0).Ref(contentid, v15 => { contentid = v15; }).Val((Int16)0).Val((Int16)0)));
                    descriptionwf = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText").Val((Int16)0).Ref(contentid, v16 => { contentid = v16; }).Val((Int16)0).Val((Int16)0)));
                    newcontentid = _.VAL(_.CALL(this, hlobj, "GenerateContentID"));
                    _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup").Val((Int16)0).Ref(newcontentid, v17 => { newcontentid = v17; }).Val((Int16)0).Ref(assignedgroupwf, v18 => { assignedgroupwf = v18; }));
                    _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson").Val((Int16)0).Ref(newcontentid, v19 => { newcontentid = v19; }).Val((Int16)0).Ref(assignedpersonwf, v20 => { assignedpersonwf = v20; }));
                    _.CALL(this, newtask, "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText").Val((Int16)0).Ref(newcontentid, v21 => { newcontentid = v21; }).Val((Int16)0).Ref(descriptionwf, v22 => { descriptionwf = v22; }));
                }
            }

            assets = _.VAL(_.CALL(this, hlobj, "GetItemsEx", _.ARGS.Val((Int16)0).Val((Int16)0).Val((Int16)131)));
            var enumerationContent3 = _.ENUMERABLE(assets).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                asset = enumerationContent3.Current;
                _.CALL(this, newtask, "AddItemEx", _.ARGS.Val((Int16)0).Ref(asset, v23 => { asset = v23; }).Val((Int16)0).Val((Int16)131));
            }
            refnumber = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("CASEINFO.REFERENCENUMBER").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.EQ(_.NullableNUM(langid), (Int16)7)))
            {
                description = _.CONCAT(description, VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine, "[Diese Aufgabe wurde automatisch durch den Systemtask mit der Bezugsnummer '", refnumber, " erstellt.]");
            }
            else
            {
                description = _.CONCAT(description, VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine, "[This Task was created automatically by Systemtask with the Reference Number '", refnumber, "'.]");
            }
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(description, v24 => { description = v24; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(exoperation, v25 => { exoperation = v25; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("CaseSpecialRouting.AssignedGroup").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(assignedgroup, v26 => { assignedgroup = v26; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("CaseSpecialRouting.AssignedPerson").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(assignedperson, v27 => { assignedperson = v27; }));
            _.CALL(this, newtask, "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(team, v28 => { team = v28; }));
            _.CALL(this, hlcontext, "SaveObject", _.ARGS.Val(newtask));
            _.CALL(this, newtask, "Unreserve", _.ARGS.ForceBrackets());
        }

        //Festlegung der Definitionen eines SystemTasks pro Tag.
        //Determining of definitions of a SystemTask by day.
        public void createsystemtaskdefbyday(ref object systaskbeginndate, ref object systaskenddate, ref object noenddate, ref object numberofdays, ref object taskdefname, ref object recurrenceendtype)
        {
            object hlobj = null;
            object hlsystemtask = null;
            object systemtaskdefinitionname = null;
            object scriptcode = null;
            object newtaskendtime = null;
            object hlsystemtaskdefinitionobj = null;
            hlobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlsystemtask = _.OBJ(_.CALL(this, _env.hlcontext, "CreateSystemTask", _.ARGS.Val((Int16)0)));
            object byrefalias4 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias4, v29 => { byrefalias4 = v29; }));
            }
            finally { taskdefname = byrefalias4; }
            systemtaskdefinitionname = _.VAL(_.CALL(this, hlsystemtask, "GetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            scriptcode = "MyTask1";
            object byrefalias5 = systaskbeginndate;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias5, v30 => { byrefalias5 = v30; }));
            }
            finally { systaskbeginndate = byrefalias5; }
            //Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
            //Check which option had been chosen in duration of the SystemTask.
            //=No EndDate
            //Alt - Anfang
            //If recurrenceEndType = "0" Then
            //	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
            //Else
            //=UserEndDate
            //	If recurrenceEndType = "2"	Then
            //		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
            //	End If
            //End If
            //Alt - Ende
            //Neu - Anfang
            if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "0")))
            {
                object byrefalias6 = systaskbeginndate;
                try
                {
                    _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias6, v31 => { byrefalias6 = v31; }));
                }
                finally { systaskbeginndate = byrefalias6; }
                recurrenceendtype = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "2")))
                {
                    object byrefalias7 = systaskenddate;
                    try
                    {
                        _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias7, v32 => { byrefalias7 = v32; }));
                    }
                    finally { systaskenddate = byrefalias7; }
                    recurrenceendtype = "1";
                }
            }
            //Neu - Ende

            object byrefalias8 = recurrenceendtype;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDTYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias8, v33 => { byrefalias8 = v33; }));
            }
            finally { recurrenceendtype = byrefalias8; }
            object byrefalias9 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias9, v34 => { byrefalias9 = v34; }));
            }
            finally { taskdefname = byrefalias9; }
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptcode, v35 => { scriptcode = v35; }));
            object byrefalias10 = numberofdays;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias10, v36 => { byrefalias10 = v36; }));
            }
            finally { numberofdays = byrefalias10; }
            _.CALL(this, _env.hlcontext, "SaveSystemTask", _.ARGS.Ref(hlsystemtask, v37 => { hlsystemtask = v37; }));
            hlsystemtaskdefinitionobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetSystemTask", _.ARGS.Ref(systemtaskdefinitionname, v38 => { systemtaskdefinitionname = v38; })));
            _.CALL(this, hlobj, "AddSystemtask", _.ARGS.Ref(hlsystemtaskdefinitionobj, v39 => { hlsystemtaskdefinitionobj = v39; }));
        }

        //Entfernt einen vorhandenen SystemTask.
        //Remove an existing SystemTask.
        public void deletesystemtask(ref object hlcontext, ref object hlobj, ref object hlsystemtask, ref object taskname)
        {
            object byrefalias11 = hlsystemtask;
            try
            {
                _.CALL(this, hlcontext, "RemoveSystemTask", _.ARGS.Ref(byrefalias11, v40 => { byrefalias11 = v40; }));
            }
            finally { hlsystemtask = byrefalias11; }
        }

        //Festlegung der Definitionen eines SystemTasks pro Woche.
        //Determining of definitions of a SystemTask by week.
        public void createsystemtaskdefbyweek(ref object systaskbeginndate, ref object systaskenddate, ref object noenddate, ref object numberofweeks, ref object mondayflag, ref object tuesdayflag, ref object wednesdayflag, ref object thursdayflag, ref object fridayflag, ref object saturdayflag, ref object sundayflag, ref object taskdefname, ref object recurrencedaymask, ref object recurrenceendtype)
        {
            object hlobj = null;
            object hlsystemtask = null;
            object systemtaskdefinitionname = null;
            object scriptcode = null;
            object hlsystemtaskdefinitionobj = null;
            hlobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlsystemtask = _.OBJ(_.CALL(this, _env.hlcontext, "CreateSystemTask", _.ARGS.Val((Int16)0)));
            object byrefalias12 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias12, v41 => { byrefalias12 = v41; }));
            }
            finally { taskdefname = byrefalias12; }
            systemtaskdefinitionname = _.VAL(_.CALL(this, hlsystemtask, "GetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            scriptcode = "MyTask1";

            //Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
            //Check which option had been chosen in duration of the SystemTask.
            //Alt - Anfang
            //If recurrenceEndType = "0" Then
            //	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
            //Else
            //	If recurrenceEndType = "2"	Then
            //		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
            //		recurrenceEndType = "1"
            //	End If
            //End If
            //Alt - Ende
            //Neu - Anfang
            if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "0")))
            {
                object byrefalias13 = systaskbeginndate;
                try
                {
                    _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias13, v42 => { byrefalias13 = v42; }));
                }
                finally { systaskbeginndate = byrefalias13; }
                recurrenceendtype = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "2")))
                {
                    object byrefalias14 = systaskenddate;
                    try
                    {
                        _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias14, v43 => { byrefalias14 = v43; }));
                    }
                    finally { systaskenddate = byrefalias14; }
                    recurrenceendtype = "1";
                }
            }
            //Neu - Ende
            object byrefalias15 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias15, v44 => { byrefalias15 = v44; }));
            }
            finally { taskdefname = byrefalias15; }
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptcode, v45 => { scriptcode = v45; }));
            object byrefalias16 = systaskbeginndate;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias16, v46 => { byrefalias16 = v46; }));
            }
            finally { systaskbeginndate = byrefalias16; }
            object byrefalias17 = recurrenceendtype;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.TYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias17, v47 => { byrefalias17 = v47; }));
            }
            finally { recurrenceendtype = byrefalias17; }
            object byrefalias18 = numberofweeks;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias18, v48 => { byrefalias18 = v48; }));
            }
            finally { numberofweeks = byrefalias18; }
            object byrefalias19 = recurrencedaymask;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.DAYMASK").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias19, v49 => { byrefalias19 = v49; }));
            }
            finally { recurrencedaymask = byrefalias19; }

            _.CALL(this, _env.hlcontext, "SaveSystemTask", _.ARGS.Ref(hlsystemtask, v50 => { hlsystemtask = v50; }));
            hlsystemtaskdefinitionobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetSystemTask", _.ARGS.Ref(systemtaskdefinitionname, v51 => { systemtaskdefinitionname = v51; })));
            _.CALL(this, hlobj, "AddSystemtask", _.ARGS.Ref(hlsystemtaskdefinitionobj, v52 => { hlsystemtaskdefinitionobj = v52; }));
        }

        //Festlegung der Definitionen eines SystemTasks pro Monat.
        //Determining of definitions of a SystemTask by month.
        public void createsystemtaskdefbymonth(ref object systaskbeginndate, ref object systaskenddate, ref object noenddate, ref object dayofmonth, ref object numberofmonths, ref object taskdefname, ref object recurrenceendtype)
        {
            object hlobj = null;
            object hlsystemtask = null;
            object systemtaskdefinitionname = null;
            object scriptcode = null;
            object hlsystemtaskdefinitionobj = null;
            hlobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlsystemtask = _.OBJ(_.CALL(this, _env.hlcontext, "CreateSystemTask", _.ARGS.Val((Int16)0)));
            object byrefalias20 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias20, v53 => { byrefalias20 = v53; }));
            }
            finally { taskdefname = byrefalias20; }
            systemtaskdefinitionname = _.VAL(_.CALL(this, hlsystemtask, "GetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            scriptcode = "MyTask1";

            //Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
            //Check which option had been chosen in duration of the SystemTask.
            //Alt - Anfang
            //If recurrenceEndType = "0" Then
            //	Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskBeginnDate)
            //Else
            //	If recurrenceEndType = "2"	Then
            //		Call hlSystemTask.SetValue("SYSTASKINFO.ENDDATE",0,0,0,SysTaskEndDate)
            //	End If
            //End If
            //Alt - Ende
            //Neu - Anfang
            if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "0")))
            {
                object byrefalias21 = systaskbeginndate;
                try
                {
                    _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias21, v54 => { byrefalias21 = v54; }));
                }
                finally { systaskbeginndate = byrefalias21; }
                recurrenceendtype = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceendtype), "2")))
                {
                    object byrefalias22 = systaskenddate;
                    try
                    {
                        _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias22, v55 => { byrefalias22 = v55; }));
                    }
                    finally { systaskenddate = byrefalias22; }
                    recurrenceendtype = "1";
                }
            }
            //Neu - Ende
            object byrefalias23 = recurrenceendtype;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.TYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias23, v56 => { byrefalias23 = v56; }));
            }
            finally { recurrenceendtype = byrefalias23; }
            object byrefalias24 = dayofmonth;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.DAYOFMONTH").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias24, v57 => { byrefalias24 = v57; }));
            }
            finally { dayofmonth = byrefalias24; }
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INSTANCE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("0"));
            object byrefalias25 = numberofmonths;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias25, v58 => { byrefalias25 = v58; }));
            }
            finally { numberofmonths = byrefalias25; }
            object byrefalias26 = systaskbeginndate;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias26, v59 => { byrefalias26 = v59; }));
            }
            finally { systaskbeginndate = byrefalias26; }
            object byrefalias27 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias27, v60 => { byrefalias27 = v60; }));
            }
            finally { taskdefname = byrefalias27; }
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptcode, v61 => { scriptcode = v61; }));
            _.CALL(this, _env.hlcontext, "SaveSystemTask", _.ARGS.Ref(hlsystemtask, v62 => { hlsystemtask = v62; }));
            hlsystemtaskdefinitionobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetSystemTask", _.ARGS.Ref(systemtaskdefinitionname, v63 => { systemtaskdefinitionname = v63; })));
            _.CALL(this, hlobj, "AddSystemtask", _.ARGS.Ref(hlsystemtaskdefinitionobj, v64 => { hlsystemtaskdefinitionobj = v64; }));
        }

        //Sub fuehrt den SystemTask einmalig aus.
        //Sub execute SystemTask one-time.
        public void createonetimesystemtask(ref object onetimetask, ref object systaskenddate, ref object systaskbeginndate, ref object taskdefname)
        {
            object hlobj = null;
            object hlsystemtask = null;
            object scriptcode = null;
            object systemtaskdefinitionname = null;
            object hlsystemtaskdefinitionobj = null;
            object errcode = null; /* Undeclared in source */
            hlobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlsystemtask = _.OBJ(_.CALL(this, _env.hlcontext, "CreateSystemTask", _.ARGS.Val("0")));
            object byrefalias28 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias28, v65 => { byrefalias28 = v65; }));
            }
            finally { taskdefname = byrefalias28; }
            scriptcode = "MyTask1";

            //Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
            //Check which option had been chosen in duration of the SystemTask.
            systemtaskdefinitionname = _.VAL(_.CALL(this, hlsystemtask, "GetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            object byrefalias29 = taskdefname;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias29, v66 => { byrefalias29 = v66; }));
            }
            finally { taskdefname = byrefalias29; }
            object byrefalias30 = systaskbeginndate;
            try
            {
                _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(byrefalias30, v67 => { byrefalias30 = v67; }));
            }
            finally { systaskbeginndate = byrefalias30; }
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val("09.09.2099 09:09:09"));
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDTYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1));
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.ENDCOUNT").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1));
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.TYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0));
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)1));
            _.CALL(this, hlsystemtask, "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptcode, v68 => { scriptcode = v68; }));

            //Wenn kein Datum angegebene wurde, muss eine Fehlermeldung angezeigt werden.
            //If no date was entered, show an error message.
            if (_.IF(_.EQ(_.NullableSTR(systaskbeginndate), "")))
            {
                errcode = "#ERR_TSKMNT_002";
            }
            _.CALL(this, _env.hlcontext, "SaveSystemTask", _.ARGS.Ref(hlsystemtask, v69 => { hlsystemtask = v69; }));
            hlsystemtaskdefinitionobj = _.OBJ(_.CALL(this, _env.hlcontext, "GetSystemTask", _.ARGS.Ref(systemtaskdefinitionname, v70 => { systemtaskdefinitionname = v70; })));
            _.CALL(this, hlobj, "AddSystemtask", _.ARGS.Ref(hlsystemtaskdefinitionobj, v71 => { hlsystemtaskdefinitionobj = v71; }));
            _.CALL(this, hlobj, "HasSystemTask", _.ARGS.Ref(systemtaskdefinitionname, v72 => { systemtaskdefinitionname = v72; }));
        }

        //----------------------------------------------------------------------------------------------------------
        public object isvalidobject(ref object obj)
        {
            return _.VAL(_.AND(_.ISOBJECT(obj), _.NOT(_.IS(obj, VBScriptConstants.Nothing))));
        }

        public void exportobject(ref object hlcontext, ref object hlobj)
        {
            object objdefname = null;
            object aliasname = null;
            object newchangeobj = null;
            object xmldoc = null;
            object nodedata = null;
            object nodeobjects = null;
            object nodeobject = null;
            object attaliasname = null;
            object nodeattributes = null;
            object noderelations = null;
            object filename = null;
            objdefname = _.VAL(_.CALL(this, hlobj, "GetType", _.ARGS.ForceBrackets()));
            aliasname = _.CONCAT("NewCI", objdefname);
            newchangeobj = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.DataToSAPAMChange").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            if (_.IF(_.OR(_.EQ(_.NullableSTR(newchangeobj), "0"), _.EQ(_.NullableSTR(newchangeobj), ""))))
            {
                aliasname = _.VAL(aliasname);
            }
            else
            {
                aliasname = _.CONCAT("ChangedCI", objdefname);
            }

            // VBScript source code
            xmldoc = _.OBJ(_.CREATEOBJECT("msxml2.DomDocument"));

            //create root element
            nodedata = _.OBJ(_.CALL(this, xmldoc, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Data")))));
            nodeobjects = _.OBJ(_.CALL(this, nodedata, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Objects")))));
            nodeobject = _.OBJ(_.CALL(this, nodeobjects, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Ref(objdefname, v73 => { objdefname = v73; })))));
            attaliasname = _.OBJ(_.CALL(this, xmldoc, "createAttribute", _.ARGS.Val("aliasname")));
            _.SET(_.VAL(aliasname), this, attaliasname, "Text");
            _.CALL(this, nodeobject, "Attributes", "setNamedItem", _.ARGS.Ref(attaliasname, v74 => { attaliasname = v74; }));
            nodeattributes = _.OBJ(_.CALL(this, nodeobject, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Attributes")))));
            noderelations = _.OBJ(_.CALL(this, nodedata, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Relations")))));
            /////////////////////////////////////////////////////////////

            ////////////////// HLOBJECT.ID
            //Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
            object byrefalias31 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias31, v75 => { byrefalias31 = v75; }).Ref(xmldoc, v76 => { xmldoc = v76; }).Ref(nodeattributes, v77 => { nodeattributes = v77; }).Val(false).Val("AssetGeneral.AssetName").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("AssetGeneral.AssetName").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias31; } // hlObj.GetValue("AssetGeneral.AssetName", 0,0,0,0)
            object byrefalias32 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias32, v78 => { byrefalias32 = v78; }).Ref(xmldoc, v79 => { xmldoc = v79; }).Ref(nodeattributes, v80 => { nodeattributes = v80; }).Val(false).Val("AccountingDetail.CostCenter").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("AccountingDetail.CostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias32; }
            object byrefalias33 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias33, v81 => { byrefalias33 = v81; }).Ref(xmldoc, v82 => { xmldoc = v82; }).Ref(nodeattributes, v83 => { nodeattributes = v83; }).Val(false).Val("AssetGeneral.Serialnumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("AssetGeneral.Serialnumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias33; }
            object byrefalias34 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias34, v84 => { byrefalias34 = v84; }).Ref(xmldoc, v85 => { xmldoc = v85; }).Ref(nodeattributes, v86 => { nodeattributes = v86; }).Val(false).Val("ProcurementDetail.AllocationNumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias34; }
            object byrefalias35 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias35, v87 => { byrefalias35 = v87; }).Ref(xmldoc, v88 => { xmldoc = v88; }).Ref(nodeattributes, v89 => { nodeattributes = v89; }).Val(false).Val("ProcurementDetail.AllocationType").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias35; }
            object byrefalias36 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias36, v90 => { byrefalias36 = v90; }).Ref(xmldoc, v91 => { xmldoc = v91; }).Ref(nodeattributes, v92 => { nodeattributes = v92; }).Val(false).Val("ProcurementDetail.OrderNumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.OrderNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias36; }
            object byrefalias37 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias37, v93 => { byrefalias37 = v93; }).Ref(xmldoc, v94 => { xmldoc = v94; }).Ref(nodeattributes, v95 => { nodeattributes = v95; }).Val(false).Val("ProcurementDetail.OrderPosition").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.OrderPosition").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias37; }
            object byrefalias38 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias38, v96 => { byrefalias38 = v96; }).Ref(xmldoc, v97 => { xmldoc = v97; }).Ref(nodeattributes, v98 => { nodeattributes = v98; }).Val(false).Val("ProcurementDetail.VendorNumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.VendorNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias38; }
            object byrefalias39 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias39, v99 => { byrefalias39 = v99; }).Ref(xmldoc, v100 => { xmldoc = v100; }).Ref(nodeattributes, v101 => { nodeattributes = v101; }).Val(true).Val("TrumpfAssetGeneral.CINumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias39; }
            object byrefalias40 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias40, v102 => { byrefalias40 = v102; }).Ref(xmldoc, v103 => { xmldoc = v103; }).Ref(nodeattributes, v104 => { nodeattributes = v104; }).Val(false).Val("TrumpfAssetGeneral.CompanyCode").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias40; }
            object byrefalias41 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias41, v105 => { byrefalias41 = v105; }).Ref(xmldoc, v106 => { xmldoc = v106; }).Ref(nodeattributes, v107 => { nodeattributes = v107; }).Val(false).Val("TrumpfAssetGeneral.InvestmentNumber").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.InvestmentNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias41; }
            object byrefalias42 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias42, v108 => { byrefalias42 = v108; }).Ref(xmldoc, v109 => { xmldoc = v109; }).Ref(nodeattributes, v110 => { nodeattributes = v110; }).Val(false).Val("TrumpfAssetGeneral.Manufacturer").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.Manufacturer").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias42; }
            object byrefalias43 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias43, v111 => { byrefalias43 = v111; }).Ref(xmldoc, v112 => { xmldoc = v112; }).Ref(nodeattributes, v113 => { nodeattributes = v113; }).Val(false).Val("TrumpfAssetGeneral.SAPCostCenter").Val(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.SAPCostCenter").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0))));
            }
            finally { hlcontext = byrefalias43; }

            // Save to File
            if (_.IF(_.OR(_.EQ(_.NullableSTR(newchangeobj), "0"), _.EQ(_.NullableSTR(newchangeobj), ""))))
            {
                filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\HELPLINE_out\\c11\\hlnew", objdefname, "_", _.CALL(this, hlobj, "GetID"), ".xml");
            }
            else
            {
                filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\HELPLINE_out\\c11\\hlchange", objdefname, "_", _.CALL(this, hlobj, "GetID"), ".xml");
            }
            _.CALL(this, xmldoc, "Save", _.ARGS.Val(filename));

        }

        public void exportobjectincident(ref object hlcontext, ref object hlobj)
        {
            object objdefname = null;
            object aliasname1 = null;
            object aliasname2 = null;
            object aliasnamesu = null;
            object elimierungsgrundde = null;
            object elimierungsgrunden = null;
            object buchungskreis = null;
            object buchungskreis1 = null;
            object teamkeyword = null;
            object kontierungsnr = null;
            object kontierungstyp = null;
            object beschreibung = null;
            object cinummer = null;
            object increqsubject = null;
            object xmldoc = null;
            object nodedata = null;
            object nodeobjects = null;
            object nodeobject = null;
            object attaliasname = null;
            object nodeattributes = null;
            object nodeserviceunits = null;
            object nodeserviceunit = null;
            object attaliasnamesu = null;
            object nodeobject2 = null;
            object attaliasname2 = null;
            object nodeattributes2 = null;
            object noderelations = null;
            object nodeproduct2case = null;
            object filename = null;
            objdefname = "IncidentRequest";
            aliasname1 = "obj1";
            aliasname2 = "obj2";
            aliasnamesu = "objSU";
            elimierungsgrundde = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)7).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            elimierungsgrunden = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetStatus.CISubStatus").Val((Int16)9).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            buchungskreis = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            buchungskreis1 = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CompanyCode").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            teamkeyword = "";
            kontierungsnr = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.AllocationNumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            kontierungstyp = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("ProcurementDetail.AllocationType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            beschreibung = "";
            beschreibung = _.CONCAT("CI ist auf Status 'Elimiert' gesetzt worden. Die CI-Nummmer steht im Betreff. Der Eliminierungsgrund lautet: ", elimierungsgrundde);
            beschreibung = _.CONCAT(beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "The CI-Status is set to Eliminated. The CI-Number is displayed in the subject of the incident. The elimination reason is: ", elimierungsgrunden);
            beschreibung = _.CONCAT(beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Kontierungsnummer: ", kontierungsnr);
            beschreibung = _.CONCAT(beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Kontierungstyp: ", kontierungstyp);
            beschreibung = _.CONCAT(beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Allocationnumber: ", kontierungsnr);
            beschreibung = _.CONCAT(beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Allocationtype: ", kontierungstyp);

            if (_.IF(_.EQ(buchungskreis, "107")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "110")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "111")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "114")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "122")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "146")))
            {
                teamkeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(buchungskreis, "222")))
            {
                teamkeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(buchungskreis, "223")))
            {
                teamkeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(buchungskreis, "225")))
            {
                teamkeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(buchungskreis, "314")))
            {
                teamkeyword = "KOControllingPasching";
            }
            else if (_.IF(_.EQ(buchungskreis, "231")))
            {
                teamkeyword = "KOControllingFarmington";
            }
            else if (_.IF(_.EQ(buchungskreis, "237")))
            {
                teamkeyword = "KOControllingCranbury";
            }

            /////////////////////////////////////////////////////////////
            cinummer = _.VAL(_.CALL(this, hlobj, "GetValue", _.ARGS.Val("TrumpfAssetGeneral.CINumber").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));
            increqsubject = _.CONCAT("Eliminierung/Elimination: ", cinummer, " Internal helpLine-ID: ", _.CALL(this, hlobj, "GetID"));

            // VBScript source code
            xmldoc = _.OBJ(_.CREATEOBJECT("msxml2.DomDocument"));

            //create root element
            nodedata = _.OBJ(_.CALL(this, xmldoc, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Data")))));
            nodeobjects = _.OBJ(_.CALL(this, nodedata, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Objects")))));

            ////// obj1: IncidentRequest///////////////////////////////////////////////////////

            nodeobject = _.OBJ(_.CALL(this, nodeobjects, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Ref(objdefname, v114 => { objdefname = v114; })))));
            attaliasname = _.OBJ(_.CALL(this, xmldoc, "createAttribute", _.ARGS.Val("aliasname")));
            _.SET(_.VAL(aliasname1), this, attaliasname, "Text");
            _.CALL(this, nodeobject, "Attributes", "setNamedItem", _.ARGS.Ref(attaliasname, v115 => { attaliasname = v115; }));
            nodeattributes = _.OBJ(_.CALL(this, nodeobject, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Attributes")))));
            nodeserviceunits = _.OBJ(_.CALL(this, nodeobject, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("ServiceUnits")))));
            nodeserviceunit = _.OBJ(_.CALL(this, nodeserviceunits, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("ServiceUnit")))));
            attaliasnamesu = _.OBJ(_.CALL(this, xmldoc, "createAttribute", _.ARGS.Val("aliasname")));
            _.SET(_.VAL(aliasnamesu), this, attaliasnamesu, "Text");
            _.CALL(this, nodeserviceunit, "Attributes", "setNamedItem", _.ARGS.Ref(attaliasnamesu, v116 => { attaliasnamesu = v116; }));
            ////////////////// HLOBJECT.ID
            //Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
            object byrefalias44 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias44, v117 => { byrefalias44 = v117; }).Ref(xmldoc, v118 => { xmldoc = v118; }).Ref(nodeattributes, v119 => { nodeattributes = v119; }).Val(true).Val("CaseGeneral.Subject").Ref(increqsubject, v120 => { increqsubject = v120; }));
            }
            finally { hlcontext = byrefalias44; }
            object byrefalias45 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias45, v121 => { byrefalias45 = v121; }).Ref(xmldoc, v122 => { xmldoc = v122; }).Ref(nodeattributes, v123 => { nodeattributes = v123; }).Val(false).Val("CaseDescription.DescriptionText").Ref(beschreibung, v124 => { beschreibung = v124; }));
            }
            finally { hlcontext = byrefalias45; }
            object byrefalias46 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias46, v125 => { byrefalias46 = v125; }).Ref(xmldoc, v126 => { xmldoc = v126; }).Ref(nodeattributes, v127 => { nodeattributes = v127; }).Val(false).Val("Keywords.KeywordOrga").Ref(teamkeyword, v128 => { teamkeyword = v128; }));
            }
            finally { hlcontext = byrefalias46; }
            object byrefalias47 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias47, v129 => { byrefalias47 = v129; }).Ref(xmldoc, v130 => { xmldoc = v130; }).Ref(nodeattributes, v131 => { nodeattributes = v131; }).Val(false).Val("CaseGeneral.CompanyCode").Ref(buchungskreis1, v132 => { buchungskreis1 = v132; }));
            }
            finally { hlcontext = byrefalias47; }
            object byrefalias48 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias48, v133 => { byrefalias48 = v133; }).Ref(xmldoc, v134 => { xmldoc = v134; }).Ref(nodeattributes, v135 => { nodeattributes = v135; }).Val(false).Val("Keywords.Keyword").Val("KWStdSWhelplineInterfaceAM"));
            }
            finally { hlcontext = byrefalias48; }
            object byrefalias49 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias49, v136 => { byrefalias49 = v136; }).Ref(xmldoc, v137 => { xmldoc = v137; }).Ref(nodeattributes, v138 => { nodeattributes = v138; }).Val(false).Val("IncidentAttribute.IncidentStatus").Val("IncidentStatusNew"));
            }
            finally { hlcontext = byrefalias49; }
            object byrefalias50 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias50, v139 => { byrefalias50 = v139; }).Ref(xmldoc, v140 => { xmldoc = v140; }).Ref(nodeattributes, v141 => { nodeattributes = v141; }).Val(false).Val("IncidentAttribute.RequestType").Val("RequestTypeService"));
            }
            finally { hlcontext = byrefalias50; }
            object byrefalias51 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias51, v142 => { byrefalias51 = v142; }).Ref(xmldoc, v143 => { xmldoc = v143; }).Ref(nodeserviceunit, v144 => { nodeserviceunit = v144; }).Val(true).Val("IncidentSUAttribute.IncidentOperation").Val("IncidentOperation"));
            }
            finally { hlcontext = byrefalias51; }

            ////// obj2: Product///////////////////////////////////////////////////////
            nodeobject2 = _.OBJ(_.CALL(this, nodeobjects, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val(_.CALL(this, hlobj, "GetType", _.ARGS.ForceBrackets()))))));
            attaliasname2 = _.OBJ(_.CALL(this, xmldoc, "createAttribute", _.ARGS.Val("aliasname")));
            _.SET(_.VAL(aliasname2), this, attaliasname2, "Text");
            _.CALL(this, nodeobject2, "Attributes", "setNamedItem", _.ARGS.Ref(attaliasname2, v145 => { attaliasname2 = v145; }));
            nodeattributes2 = _.OBJ(_.CALL(this, nodeobject2, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Attributes")))));
            object byrefalias52 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendNode", _.ARGS.Ref(byrefalias52, v146 => { byrefalias52 = v146; }).Ref(xmldoc, v147 => { xmldoc = v147; }).Ref(nodeattributes2, v148 => { nodeattributes2 = v148; }).Val(true).Val("TrumpfAssetGeneral.CINumber").Ref(cinummer, v149 => { cinummer = v149; }));
            }
            finally { hlcontext = byrefalias52; }

            ////// Relations///////////////////////////////////////////////////////
            noderelations = _.OBJ(_.CALL(this, nodedata, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Relations")))));
            nodeproduct2case = _.OBJ(_.CALL(this, noderelations, "appendChild", _.ARGS.Val(_.CALL(this, xmldoc, "createElement", _.ARGS.Val("Product2Case")))));

            object byrefalias53 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendTextNode", _.ARGS.Ref(byrefalias53, v150 => { byrefalias53 = v150; }).Ref(xmldoc, v151 => { xmldoc = v151; }).Ref(nodeproduct2case, v152 => { nodeproduct2case = v152; }).Val("Parent").Ref(aliasnamesu, v153 => { aliasnamesu = v153; }));
            }
            finally { hlcontext = byrefalias53; }
            object byrefalias54 = hlcontext;
            try
            {
                _.CALL(this, _outer, "AppendTextNode", _.ARGS.Ref(byrefalias54, v154 => { byrefalias54 = v154; }).Ref(xmldoc, v155 => { xmldoc = v155; }).Ref(nodeproduct2case, v156 => { nodeproduct2case = v156; }).Val("Child").Ref(aliasname2, v157 => { aliasname2 = v157; }));
            }
            finally { hlcontext = byrefalias54; }

            // Save to File
            filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\helpline_in\\c11\\", objdefname, "_", _.CALL(this, hlobj, "GetID"), ".xml");

            _.CALL(this, xmldoc, "Save", _.ARGS.Val(filename));

        }

        public void appendnode(ref object hlcontext, ref object xmldoc, ref object nodeobject, ref object iskey, ref object key, ref object value)
        {
            object valuenode = null;
            object cdata = null;
            object attiskey = null;
            object byrefalias55 = key;
            try
            {
                valuenode = _.OBJ(_.CALL(this, xmldoc, "createElement", _.ARGS.Ref(byrefalias55, v158 => { byrefalias55 = v158; })));
            }
            finally { key = byrefalias55; }
            object byrefalias56 = value;
            try
            {
                cdata = _.OBJ(_.CALL(this, xmldoc, "createCDATASection", _.ARGS.Ref(byrefalias56, v159 => { byrefalias56 = v159; })));
            }
            finally { value = byrefalias56; }
            _.CALL(this, valuenode, "appendChild", _.ARGS.Val(cdata));
            _.CALL(this, nodeobject, "appendChild", _.ARGS.Val(valuenode));

            attiskey = _.OBJ(_.CALL(this, xmldoc, "createAttribute", _.ARGS.Val("iskey")));
            if (_.IF(iskey))
            {
                _.SET("true", this, attiskey, "Text");
            }
            else
            {
                _.SET("false", this, attiskey, "Text");
            }
            _.CALL(this, valuenode, "Attributes", "setNamedItem", _.ARGS.Ref(attiskey, v160 => { attiskey = v160; }));

        }

        public void appendtextnode(ref object hlcontext, ref object xmldoc, ref object nodeobject, ref object key, ref object value)
        {
            object valuenode = null;
            object byrefalias57 = key;
            try
            {
                valuenode = _.OBJ(_.CALL(this, xmldoc, "createElement", _.ARGS.Ref(byrefalias57, v161 => { byrefalias57 = v161; })));
            }
            finally { key = byrefalias57; }
            _.CALL(this, nodeobject, "appendChild", _.ARGS.Val(valuenode));

            _.SET(_.VAL(value), this, valuenode, "Text");
        }

        public object dbconnectionstring(ref object hlcontext)
        {
            object DBConnectionString_retVal = null;
            object dbconnection = null;
            dbconnection = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2";

            return DBConnectionString_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlcontext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}