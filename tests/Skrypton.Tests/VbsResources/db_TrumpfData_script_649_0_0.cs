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
            HLASC_SoftwareLicenseFolderView = 110941;
            HLASC_SoftwareLicenseGroupView = 110944;
        }
        internal object HLASC_SoftwareLicenseFolderView { get; set; }
        internal object HLASC_SoftwareLicenseGroupView { get; set; }
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
        public object hlITIL2()
        {
            object hlITIL2_retVal = null;
            hlITIL2_retVal = _.CREATEOBJECT("hlStartITIL2.Global");
            _.CALLm2argp(this, _outer, "hlITIL2", "SelfCheck", _.ARGS.Ref(_env.hlContext, v => { _env.hlContext = v; }));
            return hlITIL2_retVal;
        }
        //----------------------------------------------------------------------------------------------------------
        //Deaktivieren bzw. aktivieren aller Traces, Text = Logtext im App.Log
        public void Trace(ref object hlContext, ref object text)
        {
            object text_vref = text;
            try
            {
                _.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "trace", _.ARGS.Val((Int16)1).Ref(text_vref, v2 => { text_vref = v2; }));
            }
            finally { text = text_vref; }
        }
        //----------------------------------------------------------------------------------------------------------
        //Funktion InfoMail
        //Zum Aufrufen aus EBL-Skripten von Vorgaengen
        public void InfoMail(ref object hlContext, ref object hlCase, ref object Subject, ref object MailSender, ref object Receiver, ref object CC, ref object body, ref object SendAttachments)
        {
            object Email = null;
            object AttachIDs = null;
            object AttachID = null;
            object Attachment = null;
            object MailAttachment = null;

            Email = _.OBJ(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "CreateMail"));

            //Falls der Parameter <SendAttachmnets> beim Aufruf "1" ist, werden Anhaenge mitversandt
            if (_.IF(_.EQ(_.CBOOL(SendAttachments), true)))
            {
                Attachment = VBScriptConstants.Nothing;
                AttachIDs = _.VAL(_.CALLm1v2(this, _.NnO(hlCase, "hlCase"), "GetAttachmentKeys", "HLOBJECTINFO.ATTACHMENT", (Int16)0));
                var enumerationContent = _.ENUMERABLE(AttachIDs).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent.MoveNext())
                        break;
                    AttachID = enumerationContent.Current;
                    Attachment = _.OBJ(_.CALLm1argp(this, _.NnO(hlCase, "hlCase"), "GetAttachment", _.ARGS.Val("HLOBJECTINFO.ATTACHMENT").Ref(AttachID, v3 => { AttachID = v3; }).Val((Int16)0)));
                    if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, _.NnO(Attachment, "Attachment"), "Size")), (Int16)0)))
                    {
                        MailAttachment = VBScriptConstants.Nothing;
                        MailAttachment = _.OBJ(_.CALLm1v0(this, _.NnO(Email, "Email"), "AddAttachment"));
                        _.SETm1a0(this, _.NnO(MailAttachment, "MailAttachment"), "name", _.VAL(_.CALLm1v0(this, _.NnO(Attachment, "Attachment"), "name")));
                        _.SETm1a0(this, _.NnO(MailAttachment, "MailAttachment"), "data", _.VAL(_.CALLm1v0(this, _.NnO(Attachment, "Attachment"), "data")));
                    }
                }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(MailSender), "")))
            {
                _.SETm1a0(this, _.NnO(Email, "Email"), "SenderMail", _.VAL(MailSender));
            }
            _.SETm1a0(this, _.NnO(Email, "Email"), "To", _.VAL(Receiver));
            _.SETm1a0(this, _.NnO(Email, "Email"), "Subject", _.VAL(Subject));
            _.SETm1a0(this, _.NnO(Email, "Email"), "Body", _.VAL(body));
            if (_.IF(_.NOTEQ(_.NullableSTR(CC), "")))
            {
                _.SETm1a0(this, _.NnO(Email, "Email"), "CC", _.VAL(CC));
            }
            _.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "SendRequestMail", _.ARGS.Ref(Email, v4 => { Email = v4; }));
        }
        //----------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------
        public void CreateSubject(ref object hlContext, ref object Survey, ref object hlCaller)
        {
            object language = null;
            language = _.VAL(_.CALLm1v5(this, _.NnO(hlCaller, "hlCaller"), "GetValue", "PersonGeneral.Language", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            if (_.IF(_.EQ(_.NullableSTR(language), "LanguageGerman")))
            {
                _.CALLm1v5(this, _.NnO(Survey, "Survey"), "SetValue", "CaseGeneral.Subject", (Int16)0, (Int16)0, (Int16)0, "Umfrage zur Service-Leistung ihres Support-Teams");
            }
            else
            {
                _.CALLm1v5(this, _.NnO(Survey, "Survey"), "SetValue", "CaseGeneral.Subject", (Int16)0, (Int16)0, (Int16)0, "Survey about the Service-Quality from your Support-Team");
            }
        }
        //----------------------------------------------------------------------------------------------------------
        public void InviteSurveyEmail(ref object hlContext, ref object hlCase, ref object hlCaller)
        {
            object SUIDx = null;
            object MailRequest = null;
            object strCRLF = null;
            object Creationdate = null;
            object Datum = null;
            object Subject = null;
            object body = null;
            object refnumber = null;
            object portallink = null;
            object surname = null;
            object letteraddress = null;
            object Anrede = null;
            object PersonAddress = null;
            object language = null;
            object Email = null;
            object Emailadress = null;
            //Email an den Anfrager eines Survey-Vorgangs, um diesen zur Teilnahme an der
            //Umfrage aufzufordern.
            //Email to Requester of a Survey-Case to invite him to take part on the survey
            object hlCase_vref = hlCase, hlContext_vref = hlContext;
            try
            {
                SUIDx = _.VAL(_.CALLm2argp(this, _outer, "hlITIL2", "GetLastSUIdx", _.ARGS.Ref(hlCase_vref, v5 => { hlCase_vref = v5; }).Ref(hlContext_vref, v6 => { hlContext_vref = v6; })));
            }
            finally { hlCase = hlCase_vref; hlContext = hlContext_vref; }
            MailRequest = _.VAL(_.CALLm1v5(this, _.NnO(hlCase, "hlCase"), "GetValue", "CaseGeneral.DefaultNotification", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            if (_.IF(_.ANDe2(_.CBOOL(_.EQ(_.NullableSTR(MailRequest), "DefaultNotificationEmail")) && _.CBOOL(_.EQ(_.NullableNUM(SUIDx), (Int16)1)))))
            {
                strCRLF = _.CONCAT(_.CHR((Int16)13), _.CHR((Int16)10));
                refnumber = _.VAL(_.CALLm1v5(this, _.NnO(hlCase, "hlCase"), "GetValue", "CASEINFO.REFERENCENUMBER", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                portallink = "http://localhost/helplineportal/";
                surname = _.VAL(_.CALLm1v5(this, _.NnO(hlCaller, "hlCaller"), "GetValue", "PersonGeneral.PersonSurname", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                letteraddress = _.VAL(_.CALLm1v5(this, _.NnO(hlCaller, "hlCaller"), "GetValue", "PersonGeneral.ShortLetterAddress", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                Anrede = "Sehr geehrte Damen und Herren,";
                PersonAddress = "Dear Mrs./Ms. or Mr.,";
                language = _.VAL(_.CALLm1v5(this, _.NnO(hlCaller, "hlCaller"), "GetValue", "PersonGeneral.Language", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

                if (_.IF(_.EQ(_.NullableSTR(language), "LanguageGerman")))
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Herr/Frau";
                    }
                    Anrede = _.CONCAT("Sehr geehrte(r) ", _.CSTR(letteraddress), " ", _.CSTR(surname), ",");

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    Creationdate = _.VAL(_.CALLm1v5(this, _.NnO(hlCase, "hlCase"), "GetValue", "HLOBJECTINFO.CREATIONTIME", (Int16)7, (Int16)0, (Int16)0, (Int16)0));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    Subject = "Umfrage zur Service-Leistung ihres Support-Teams";

                    //Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
                    //If the case was recorded, the requester receives the according information
                    body = _.CSTR(Anrede);
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Wir wollen besser werden!");
                    body = _.CONCAT(body, strCRLF, "Dazu benoetigen wir Ihre Unterstuetzung und Ihr Feedback.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Sie wurden am ", Datum, " durch ein Zufallsverfahren ausgewaehlt, an einer Umfrage zu den Service-Leistungen Ihres Support-Teams teilzunehmen.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Die Teilnahme ist freiwillig und erfolgt ueber das helpLine Portal.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Rufen Sie im Browser bitte folgende URL auf:");
                    body = _.CONCAT(body, strCRLF, portallink, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Klicken Sie unter 'Ihre Anfragen' auf den Eintrag 'Umfragen'. ");
                    body = _.CONCAT(body, "Dort finden Sie das Umfrage-Formular mit der Nummer ", refnumber, ". ");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Wir freuen uns sehr, wenn Sie sich die Zeit nehmen, die Fragen zu beantworten.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Wir bedanken uns fuer Ihre Unterstuetzung!");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, strCRLF, "Mit freundlichen Gruessen");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Ihr Support Team");
                }
                else
                {
                    if (_.IF(_.EQ(_.NullableSTR(letteraddress), "")))
                    {
                        letteraddress = "Mrs./Ms./Mr.";
                    }
                    PersonAddress = _.CONCAT("Dear ", _.ADD(_.CSTR(letteraddress), " "), _.CSTR(surname), ",");

                    //Hier wird die Betreffzeile erstellt
                    //The subject field is entered here
                    Creationdate = _.VAL(_.CALLm1v5(this, _.NnO(hlCase, "hlCase"), "GetValue", "HLOBJECTINFO.CREATIONTIME", (Int16)7, (Int16)0, (Int16)0, (Int16)0));
                    Datum = _.VAL(_.MID(Creationdate, (Int16)1, (Int16)10));
                    Subject = "Survey about the Service-Quality from your Support-Team";

                    //Wenn der Vorgang aufgenommen wurde erhaelt der Anfrager darueber eine Information
                    //If the case was recorded, the requester receives the according information
                    body = _.CSTR(PersonAddress);
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "We would like to improve the efficiency of Service-Support!");
                    body = _.CONCAT(body, strCRLF, "Therefore we need your assistance and your feedback.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "You where chosen by random on ", Datum, " to take part on the Survey.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "The participation on the survey is voluntarily. You can take part on the survey via the helpLine Portal.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Start your Browser and choose the following URL:");
                    body = _.CONCAT(body, strCRLF, portallink, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Then klick 'Survey' in the menue 'Your Requests'. ", strCRLF);
                    body = _.CONCAT(body, "There, you will find the Questionnaire with the reference number ", refnumber, ". ");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "It would be nice, if you invest your time to response the questions.");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "We thank you for your assistance!");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, strCRLF, "With best regards");
                    body = _.CONCAT(body, strCRLF, strCRLF);
                    body = _.CONCAT(body, "Yours Support Team");
                }

                Email = _.OBJ(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "CreateMail"));

                //Ermittle die Emailadresse des Anfragers
                //Detect email adress of requester
                Emailadress = _.VAL(_.CALLm1v5(this, _.NnO(hlCaller, "hlCaller"), "GetValue", "PersonInformation.EmailAddress", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                if (_.IF(_.EQ(_.NullableSTR(Emailadress), "")))
                {
                    Emailadress = "Username@yourcompany.com";
                    Subject = "Diese EMail konnte nicht zugestellt werden";
                    body = "Die Mail fuer die Anfragenummer ";
                    body = _.CONCAT(body, _.CALLm1v5(this, _.NnO(hlCase, "hlCase"), "GetValue", "CASEINFO.REFERENCENUMBER", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                    body = _.CONCAT(body, " konnte wegen einer fehlenden E-Mail Adresse nicht zugestellt werden.");
                }
                _.SETm1a0(this, _.NnO(Email, "Email"), "To", _.VAL(Emailadress));
                _.SETm1a0(this, _.NnO(Email, "Email"), "Subject", _.VAL(Subject));
                _.SETm1a0(this, _.NnO(Email, "Email"), "Body", _.VAL(body));
                _.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "SendRequestMail", _.ARGS.Ref(Email, v7 => { Email = v7; }));
            }
        }
        //----------------------------------------------------------------------------------------------------------
        //Diese Funktion steuert den SystemTask wenn dieser im Vorgangstyp Task konfiguriert wurde.
        //This function controls a SystemTask if it had been configured within the casetype Task.
        public void MyTask1(ref object hlContext)
        {
            object hlObj = null;
            object lcid = null;
            object LangID = null;
            object Priority = null;
            object TaskType = null;
            object Subject = null;
            object Description = null;
            object ExOperation = null;
            object AssignedGroup = null;
            object AssignedPerson = null;
            object Team = null;
            object newTask = null;
            object hasContent = null;
            object contentIDs = null;
            object contentID = null;
            object newContentID = null;
            object assignedGroupWF = null;
            object assignedPersonWF = null;
            object descriptionWF = null;
            object subjectWF = null;
            object Assets = null;
            object Asset = null;
            object refnumber = null;
            hlObj = _.OBJ(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "GetCurrentObject", _.ARGS.ForceBrackets()));
            lcid = (Int16)0;
            lcid = _.VAL(_.CALLm1v0(this, _.NnO(hlContext, "hlContext"), "GetLocaleID"));
            LangID = (Int16)0;
            LangID = _.VAL(_.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "LangIDFromLCID", _.ARGS.Ref(lcid, v8 => { lcid = v8; })));

            //Gesetzte Daten aus dem aktuellen Task auslesen, diese werden dem zu erzeugenden Systemtask mitgegeben.
            //Read setted data of current task and take them into the created Systemtask.
            Priority = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CaseClassificationAttribute.Priority", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            TaskType = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TaskGeneral.TaskType", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Subject = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TaskGeneral.Subject", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Description = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CaseDescription.DescriptionText", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            ExOperation = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CaseDiagnosis.DiagnosisText", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            AssignedGroup = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CaseSpecialRouting.AssignedGroup", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            AssignedPerson = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CaseSpecialRouting.AssignedPerson", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Team = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "Keywords.KeywordOrga", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            newTask = _.OBJ(_.CALLm1v1(this, _.NnO(hlContext, "hlContext"), "createobject", "Task"));

            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("CaseClassificationAttribute.Priority").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(Priority, v9 => { Priority = v9; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskGeneral.TaskType").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(TaskType, v10 => { TaskType = v10; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskGeneral.Subject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(Subject, v11 => { Subject = v11; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(Team, v12 => { Team = v12; }));

            hasContent = _.VAL(_.CALLm1v3(this, _.NnO(hlObj, "hlObj"), "HasContent", "TaskDesignWorkflow.TaskWorkflow_CA", (Int16)0, (Int16)0));
            if (_.IF(_.NOTEQ(_.NullableNUM(hasContent), (Int16)0)))
            {
                subjectWF = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TaskDesignWorkflow.FlagWorkflowSubject", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
                _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskDesignWorkflow.FlagWorkflowSubject").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(subjectWF, v13 => { subjectWF = v13; }));
                contentIDs = _.VAL(_.CALLm1v2(this, _.NnO(hlObj, "hlObj"), "GetContentIDs", "TaskDesignWorkflow.TaskWorkflow_CA", (Int16)0));
                _.CALLm1v5(this, _.NnO(newTask, "newTask"), "SetValue", "TaskWorkflowAttribute.WorkflowStep", (Int16)0, (Int16)0, (Int16)0, (Int16)1);
                var enumerationContent2 = _.ENUMERABLE(contentIDs).GetEnumerator();
                while (true)
                {
                    if (!enumerationContent2.MoveNext())
                        break;
                    contentID = enumerationContent2.Current;
                    assignedGroupWF = _.VAL(_.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup").Val((Int16)0).Ref(contentID, v14 => { contentID = v14; }).Val((Int16)0).Val((Int16)0)));
                    assignedPersonWF = _.VAL(_.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson").Val((Int16)0).Ref(contentID, v15 => { contentID = v15; }).Val((Int16)0).Val((Int16)0)));
                    descriptionWF = _.VAL(_.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "GetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText").Val((Int16)0).Ref(contentID, v16 => { contentID = v16; }).Val((Int16)0).Val((Int16)0)));
                    newContentID = _.VAL(_.CALLm1v0(this, _.NnO(hlObj, "hlObj"), "GenerateContentID"));
                    _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedGroup").Val((Int16)0).Ref(newContentID, v17 => { newContentID = v17; }).Val((Int16)0).Ref(assignedGroupWF, v18 => { assignedGroupWF = v18; }));
                    _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.AssignedPerson").Val((Int16)0).Ref(newContentID, v19 => { newContentID = v19; }).Val((Int16)0).Ref(assignedPersonWF, v20 => { assignedPersonWF = v20; }));
                    _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("TaskDesignWorkflow.TaskWorkflow_CA.DescriptionText").Val((Int16)0).Ref(newContentID, v21 => { newContentID = v21; }).Val((Int16)0).Ref(descriptionWF, v22 => { descriptionWF = v22; }));
                }
            }

            Assets = _.VAL(_.CALLm1v3(this, _.NnO(hlObj, "hlObj"), "GetItemsEx", (Int16)0, (Int16)0, (Int16)131));
            var enumerationContent3 = _.ENUMERABLE(Assets).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                Asset = enumerationContent3.Current;
                _.CALLm1argp(this, _.NnO(newTask, "newTask"), "AddItemEx", _.ARGS.Val((Int16)0).Ref(Asset, v23 => { Asset = v23; }).Val((Int16)0).Val((Int16)131));
            }
            refnumber = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "CASEINFO.REFERENCENUMBER", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            if (_.IF(_.EQ(_.NullableNUM(LangID), (Int16)7)))
            {
                Description = _.CONCAT(Description, VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine, "[Diese Aufgabe wurde automatisch durch den Systemtask mit der Bezugsnummer '", refnumber, " erstellt.]");
            }
            else
            {
                Description = _.CONCAT(Description, VBScriptConstants.vbNewLine, VBScriptConstants.vbNewLine, "[This Task was created automatically by Systemtask with the Reference Number '", refnumber, "'.]");
            }
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("CaseDescription.DescriptionText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(Description, v24 => { Description = v24; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("CaseDiagnosis.DiagnosisText").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(ExOperation, v25 => { ExOperation = v25; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("CaseSpecialRouting.AssignedGroup").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(AssignedGroup, v26 => { AssignedGroup = v26; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("CaseSpecialRouting.AssignedPerson").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(AssignedPerson, v27 => { AssignedPerson = v27; }));
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "SetValue", _.ARGS.Val("Keywords.KeywordOrga").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(Team, v28 => { Team = v28; }));
            _.CALLm1v1(this, _.NnO(hlContext, "hlContext"), "SaveObject", newTask);
            _.CALLm1argp(this, _.NnO(newTask, "newTask"), "Unreserve", _.ARGS.ForceBrackets());
        }
        //Festlegung der Definitionen eines SystemTasks pro Tag.
        //Determining of definitions of a SystemTask by day.
        public void CreateSystemTaskDefbyDay(ref object SysTaskBeginnDate, ref object SysTaskEndDate, ref object NoEndDate, ref object NumberOfDays, ref object taskDefname, ref object recurrenceEndType)
        {
            object hlObj = null;
            object hlSystemTask = null;
            object systemTaskDefinitionName = null;
            object scriptCode = null;
            object newTaskEndTime = null;
            object hlSystemTaskDefinitionObj = null;
            hlObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlSystemTask = _.OBJ(_.CALLm1v1(this, _.NnO(_env.hlContext, "hlContext"), "CreateSystemTask", (Int16)0));
            object taskDefname_vref = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref, v29 => { taskDefname_vref = v29; }));
            }
            finally { taskDefname = taskDefname_vref; }
            systemTaskDefinitionName = _.VAL(_.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "GetValue", "SYSTASKINFO.DEFNAME", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            scriptCode = "MyTask1";
            object SysTaskBeginnDate_vref = SysTaskBeginnDate;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref, v30 => { SysTaskBeginnDate_vref = v30; }));
            }
            finally { SysTaskBeginnDate = SysTaskBeginnDate_vref; }
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
            if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "0")))
            {
                object SysTaskBeginnDate_vref2 = SysTaskBeginnDate;
                try
                {
                    _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref2, v31 => { SysTaskBeginnDate_vref2 = v31; }));
                }
                finally { SysTaskBeginnDate = SysTaskBeginnDate_vref2; }
                recurrenceEndType = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "2")))
                {
                    object SysTaskEndDate_vref = SysTaskEndDate;
                    try
                    {
                        _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskEndDate_vref, v32 => { SysTaskEndDate_vref = v32; }));
                    }
                    finally { SysTaskEndDate = SysTaskEndDate_vref; }
                    recurrenceEndType = "1";
                }
            }
            //Neu - Ende

            object recurrenceEndType_vref = recurrenceEndType;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDTYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(recurrenceEndType_vref, v33 => { recurrenceEndType_vref = v33; }));
            }
            finally { recurrenceEndType = recurrenceEndType_vref; }
            object taskDefname_vref2 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref2, v34 => { taskDefname_vref2 = v34; }));
            }
            finally { taskDefname = taskDefname_vref2; }
            _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptCode, v35 => { scriptCode = v35; }));
            object NumberOfDays_vref = NumberOfDays;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(NumberOfDays_vref, v36 => { NumberOfDays_vref = v36; }));
            }
            finally { NumberOfDays = NumberOfDays_vref; }
            _.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "SaveSystemTask", _.ARGS.Ref(hlSystemTask, v37 => { hlSystemTask = v37; }));
            hlSystemTaskDefinitionObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetSystemTask", _.ARGS.Ref(systemTaskDefinitionName, v38 => { systemTaskDefinitionName = v38; })));
            _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "AddSystemtask", _.ARGS.Ref(hlSystemTaskDefinitionObj, v39 => { hlSystemTaskDefinitionObj = v39; }));
        }
        //Entfernt einen vorhandenen SystemTask.
        //Remove an existing SystemTask.
        public void DeleteSystemTask(ref object hlContext, ref object hlObj, ref object hlSystemTask, ref object taskname)
        {
            object hlSystemTask_vref = hlSystemTask;
            try
            {
                _.CALLm1argp(this, _.NnO(hlContext, "hlContext"), "RemoveSystemTask", _.ARGS.Ref(hlSystemTask_vref, v40 => { hlSystemTask_vref = v40; }));
            }
            finally { hlSystemTask = hlSystemTask_vref; }
        }
        //Festlegung der Definitionen eines SystemTasks pro Woche.
        //Determining of definitions of a SystemTask by week.
        public void CreateSystemTaskDefbyWeek(ref object SysTaskBeginnDate, ref object SysTaskEndDate, ref object NoEndDate, ref object NumberOfWeeks, ref object MondayFlag, ref object TuesdayFlag, ref object WednesdayFlag, ref object ThursdayFlag, ref object FridayFlag, ref object SaturdayFlag, ref object SundayFlag, ref object taskDefname, ref object recurrencedaymask, ref object recurrenceEndType)
        {
            object hlObj = null;
            object hlSystemTask = null;
            object systemTaskDefinitionName = null;
            object scriptCode = null;
            object hlSystemTaskDefinitionObj = null;
            hlObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlSystemTask = _.OBJ(_.CALLm1v1(this, _.NnO(_env.hlContext, "hlContext"), "CreateSystemTask", (Int16)0));
            object taskDefname_vref3 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref3, v41 => { taskDefname_vref3 = v41; }));
            }
            finally { taskDefname = taskDefname_vref3; }
            systemTaskDefinitionName = _.VAL(_.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "GetValue", "SYSTASKINFO.DEFNAME", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            scriptCode = "MyTask1";

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
            if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "0")))
            {
                object SysTaskBeginnDate_vref3 = SysTaskBeginnDate;
                try
                {
                    _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref3, v42 => { SysTaskBeginnDate_vref3 = v42; }));
                }
                finally { SysTaskBeginnDate = SysTaskBeginnDate_vref3; }
                recurrenceEndType = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "2")))
                {
                    object SysTaskEndDate_vref2 = SysTaskEndDate;
                    try
                    {
                        _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskEndDate_vref2, v43 => { SysTaskEndDate_vref2 = v43; }));
                    }
                    finally { SysTaskEndDate = SysTaskEndDate_vref2; }
                    recurrenceEndType = "1";
                }
            }
            //Neu - Ende
            object taskDefname_vref4 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref4, v44 => { taskDefname_vref4 = v44; }));
            }
            finally { taskDefname = taskDefname_vref4; }
            _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptCode, v45 => { scriptCode = v45; }));
            object SysTaskBeginnDate_vref4 = SysTaskBeginnDate;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref4, v46 => { SysTaskBeginnDate_vref4 = v46; }));
            }
            finally { SysTaskBeginnDate = SysTaskBeginnDate_vref4; }
            object recurrenceEndType_vref2 = recurrenceEndType;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.TYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(recurrenceEndType_vref2, v47 => { recurrenceEndType_vref2 = v47; }));
            }
            finally { recurrenceEndType = recurrenceEndType_vref2; }
            object NumberOfWeeks_vref = NumberOfWeeks;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(NumberOfWeeks_vref, v48 => { NumberOfWeeks_vref = v48; }));
            }
            finally { NumberOfWeeks = NumberOfWeeks_vref; }
            object recurrencedaymask_vref = recurrencedaymask;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.DAYMASK").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(recurrencedaymask_vref, v49 => { recurrencedaymask_vref = v49; }));
            }
            finally { recurrencedaymask = recurrencedaymask_vref; }

            _.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "SaveSystemTask", _.ARGS.Ref(hlSystemTask, v50 => { hlSystemTask = v50; }));
            hlSystemTaskDefinitionObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetSystemTask", _.ARGS.Ref(systemTaskDefinitionName, v51 => { systemTaskDefinitionName = v51; })));
            _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "AddSystemtask", _.ARGS.Ref(hlSystemTaskDefinitionObj, v52 => { hlSystemTaskDefinitionObj = v52; }));
        }
        //Festlegung der Definitionen eines SystemTasks pro Monat.
        //Determining of definitions of a SystemTask by month.
        public void CreateSystemTaskDefbyMonth(ref object SysTaskBeginnDate, ref object SysTaskEndDate, ref object NoEndDate, ref object DayOfMonth, ref object NumberOfMonths, ref object taskDefname, ref object recurrenceEndType)
        {
            object hlObj = null;
            object hlSystemTask = null;
            object systemTaskDefinitionName = null;
            object scriptCode = null;
            object hlSystemTaskDefinitionObj = null;
            hlObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlSystemTask = _.OBJ(_.CALLm1v1(this, _.NnO(_env.hlContext, "hlContext"), "CreateSystemTask", (Int16)0));
            object taskDefname_vref5 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref5, v53 => { taskDefname_vref5 = v53; }));
            }
            finally { taskDefname = taskDefname_vref5; }
            systemTaskDefinitionName = _.VAL(_.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "GetValue", "SYSTASKINFO.DEFNAME", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            scriptCode = "MyTask1";

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
            if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "0")))
            {
                object SysTaskBeginnDate_vref5 = SysTaskBeginnDate;
                try
                {
                    _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref5, v54 => { SysTaskBeginnDate_vref5 = v54; }));
                }
                finally { SysTaskBeginnDate = SysTaskBeginnDate_vref5; }
                recurrenceEndType = "1";
            }
            else
            {
                if (_.IF(_.EQ(_.NullableSTR(recurrenceEndType), "2")))
                {
                    object SysTaskEndDate_vref3 = SysTaskEndDate;
                    try
                    {
                        _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.ENDDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskEndDate_vref3, v55 => { SysTaskEndDate_vref3 = v55; }));
                    }
                    finally { SysTaskEndDate = SysTaskEndDate_vref3; }
                    recurrenceEndType = "1";
                }
            }
            //Neu - Ende
            object recurrenceEndType_vref3 = recurrenceEndType;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.TYPE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(recurrenceEndType_vref3, v56 => { recurrenceEndType_vref3 = v56; }));
            }
            finally { recurrenceEndType = recurrenceEndType_vref3; }
            object DayOfMonth_vref = DayOfMonth;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.DAYOFMONTH").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(DayOfMonth_vref, v57 => { DayOfMonth_vref = v57; }));
            }
            finally { DayOfMonth = DayOfMonth_vref; }
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.RECURRENCE.INSTANCE", (Int16)0, (Int16)0, (Int16)0, "0");
            object NumberOfMonths_vref = NumberOfMonths;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.RECURRENCE.INTERVAL").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(NumberOfMonths_vref, v58 => { NumberOfMonths_vref = v58; }));
            }
            finally { NumberOfMonths = NumberOfMonths_vref; }
            object SysTaskBeginnDate_vref6 = SysTaskBeginnDate;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref6, v59 => { SysTaskBeginnDate_vref6 = v59; }));
            }
            finally { SysTaskBeginnDate = SysTaskBeginnDate_vref6; }
            object taskDefname_vref6 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref6, v60 => { taskDefname_vref6 = v60; }));
            }
            finally { taskDefname = taskDefname_vref6; }
            _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptCode, v61 => { scriptCode = v61; }));
            _.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "SaveSystemTask", _.ARGS.Ref(hlSystemTask, v62 => { hlSystemTask = v62; }));
            hlSystemTaskDefinitionObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetSystemTask", _.ARGS.Ref(systemTaskDefinitionName, v63 => { systemTaskDefinitionName = v63; })));
            _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "AddSystemtask", _.ARGS.Ref(hlSystemTaskDefinitionObj, v64 => { hlSystemTaskDefinitionObj = v64; }));
        }
        //Sub fuehrt den SystemTask einmalig aus.
        //Sub execute SystemTask one-time.
        public void CreateOneTimeSystemTask(ref object OneTimeTask, ref object SysTaskEndDate, ref object SysTaskBeginnDate, ref object taskDefname)
        {
            object hlObj = null;
            object hlSystemTask = null;
            object scriptCode = null;
            object systemTaskDefinitionName = null;
            object hlSystemTaskDefinitionObj = null;
            object errCode = null; /* Undeclared in source */
            hlObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetCurrentObject", _.ARGS.ForceBrackets()));
            hlSystemTask = _.OBJ(_.CALLm1v1(this, _.NnO(_env.hlContext, "hlContext"), "CreateSystemTask", "0"));
            object taskDefname_vref7 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref7, v65 => { taskDefname_vref7 = v65; }));
            }
            finally { taskDefname = taskDefname_vref7; }
            scriptCode = "MyTask1";

            //Prueft welche Option zu Duration des SystemTasks ausgewaehlt wurde.
            //Check which option had been chosen in duration of the SystemTask.
            systemTaskDefinitionName = _.VAL(_.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "GetValue", "SYSTASKINFO.DEFNAME", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            object taskDefname_vref8 = taskDefname;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.DEFNAME").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(taskDefname_vref8, v66 => { taskDefname_vref8 = v66; }));
            }
            finally { taskDefname = taskDefname_vref8; }
            object SysTaskBeginnDate_vref7 = SysTaskBeginnDate;
            try
            {
                _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.STARTDATE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(SysTaskBeginnDate_vref7, v67 => { SysTaskBeginnDate_vref7 = v67; }));
            }
            finally { SysTaskBeginnDate = SysTaskBeginnDate_vref7; }
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.ENDDATE", (Int16)0, (Int16)0, (Int16)0, "09.09.2099 09:09:09");
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.ENDTYPE", (Int16)0, (Int16)0, (Int16)0, (Int16)1);
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.ENDCOUNT", (Int16)0, (Int16)0, (Int16)0, (Int16)1);
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.RECURRENCE.TYPE", (Int16)0, (Int16)0, (Int16)0, (Int16)0);
            _.CALLm1v5(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", "SYSTASKINFO.RECURRENCE.INTERVAL", (Int16)0, (Int16)0, (Int16)0, (Int16)1);
            _.CALLm1argp(this, _.NnO(hlSystemTask, "hlSystemTask"), "SetValue", _.ARGS.Val("SYSTASKINFO.SCRIPTCODE").Val((Int16)0).Val((Int16)0).Val((Int16)0).Ref(scriptCode, v68 => { scriptCode = v68; }));

            //Wenn kein Datum angegebene wurde, muss eine Fehlermeldung angezeigt werden.
            //If no date was entered, show an error message.
            if (_.IF(_.EQ(_.NullableSTR(SysTaskBeginnDate), "")))
            {
                errCode = "#ERR_TSKMNT_002";
            }
            _.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "SaveSystemTask", _.ARGS.Ref(hlSystemTask, v69 => { hlSystemTask = v69; }));
            hlSystemTaskDefinitionObj = _.OBJ(_.CALLm1argp(this, _.NnO(_env.hlContext, "hlContext"), "GetSystemTask", _.ARGS.Ref(systemTaskDefinitionName, v70 => { systemTaskDefinitionName = v70; })));
            _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "AddSystemtask", _.ARGS.Ref(hlSystemTaskDefinitionObj, v71 => { hlSystemTaskDefinitionObj = v71; }));
            _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "HasSystemTask", _.ARGS.Ref(systemTaskDefinitionName, v72 => { systemTaskDefinitionName = v72; }));
        }
        //----------------------------------------------------------------------------------------------------------
        public object IsValidObject(ref object obj)
        {
            return _.VAL(_.ANDe2(_.CBOOL(_.ISOBJECT(obj)) && _.CBOOL(_.NOT(_.IS(obj, VBScriptConstants.Nothing)))));
        }
        public void ExportObject(ref object hlContext, ref object hlObj)
        {
            object objDefname = null;
            object aliasname = null;
            object NewChangeObj = null;
            object xmldoc = null;
            object nodeData = null;
            object nodeObjects = null;
            object nodeObject = null;
            object attAliasName = null;
            object nodeAttributes = null;
            object nodeRelations = null;
            object Filename = null;
            objDefname = _.VAL(_.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "GetType", _.ARGS.ForceBrackets()));
            aliasname = _.CONCAT("NewCI", objDefname);
            NewChangeObj = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.DataToSAPAMChange", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            if (_.IF(_.OR(_.EQ(_.NullableSTR(NewChangeObj), "0"), _.EQ(_.NullableSTR(NewChangeObj), ""))))
            {
                aliasname = _.VAL(aliasname);
            }
            else
            {
                aliasname = _.CONCAT("ChangedCI", objDefname);
            }

            // VBScript source code
            xmldoc = _.CREATEOBJECT("msxml2.DomDocument");

            //create root element
            nodeData = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Data")));
            nodeObjects = _.OBJ(_.CALLm1v1(this, _.NnO(nodeData, "nodeData"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Objects")));
            nodeObject = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObjects, "nodeObjects"), "appendChild", _.CALLm1argp(this, _.NnO(xmldoc, "xmldoc"), "createElement", _.ARGS.Ref(objDefname, v73 => { objDefname = v73; }))));
            attAliasName = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createAttribute", "aliasname"));
            _.SETm1a0(this, _.NnO(attAliasName, "attAliasName"), "Text", _.VAL(aliasname));
            _.CALLm2argp(this, _.NnO(nodeObject, "nodeObject"), "Attributes", "setNamedItem", _.ARGS.Ref(attAliasName, v74 => { attAliasName = v74; }));
            nodeAttributes = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObject, "nodeObject"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Attributes")));
            nodeRelations = _.OBJ(_.CALLm1v1(this, _.NnO(nodeData, "nodeData"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Relations")));
            /////////////////////////////////////////////////////////////

            ////////////////// HLOBJECT.ID
            //Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
            object hlContext_vref2 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref2, v75 => { hlContext_vref2 = v75; }).Ref(xmldoc, v76 => { xmldoc = v76; }).Ref(nodeAttributes, v77 => { nodeAttributes = v77; }).Val(false).Val("AssetGeneral.AssetName").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "AssetGeneral.AssetName", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref2; } // hlObj.GetValue("AssetGeneral.AssetName", 0,0,0,0)
            object hlContext_vref3 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref3, v78 => { hlContext_vref3 = v78; }).Ref(xmldoc, v79 => { xmldoc = v79; }).Ref(nodeAttributes, v80 => { nodeAttributes = v80; }).Val(false).Val("AccountingDetail.CostCenter").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "AccountingDetail.CostCenter", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref3; }
            object hlContext_vref4 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref4, v81 => { hlContext_vref4 = v81; }).Ref(xmldoc, v82 => { xmldoc = v82; }).Ref(nodeAttributes, v83 => { nodeAttributes = v83; }).Val(false).Val("AssetGeneral.Serialnumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "AssetGeneral.Serialnumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref4; }
            object hlContext_vref5 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref5, v84 => { hlContext_vref5 = v84; }).Ref(xmldoc, v85 => { xmldoc = v85; }).Ref(nodeAttributes, v86 => { nodeAttributes = v86; }).Val(false).Val("ProcurementDetail.AllocationNumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.AllocationNumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref5; }
            object hlContext_vref6 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref6, v87 => { hlContext_vref6 = v87; }).Ref(xmldoc, v88 => { xmldoc = v88; }).Ref(nodeAttributes, v89 => { nodeAttributes = v89; }).Val(false).Val("ProcurementDetail.AllocationType").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.AllocationType", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref6; }
            object hlContext_vref7 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref7, v90 => { hlContext_vref7 = v90; }).Ref(xmldoc, v91 => { xmldoc = v91; }).Ref(nodeAttributes, v92 => { nodeAttributes = v92; }).Val(false).Val("ProcurementDetail.OrderNumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.OrderNumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref7; }
            object hlContext_vref8 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref8, v93 => { hlContext_vref8 = v93; }).Ref(xmldoc, v94 => { xmldoc = v94; }).Ref(nodeAttributes, v95 => { nodeAttributes = v95; }).Val(false).Val("ProcurementDetail.OrderPosition").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.OrderPosition", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref8; }
            object hlContext_vref9 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref9, v96 => { hlContext_vref9 = v96; }).Ref(xmldoc, v97 => { xmldoc = v97; }).Ref(nodeAttributes, v98 => { nodeAttributes = v98; }).Val(false).Val("ProcurementDetail.VendorNumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.VendorNumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref9; }
            object hlContext_vref10 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref10, v99 => { hlContext_vref10 = v99; }).Ref(xmldoc, v100 => { xmldoc = v100; }).Ref(nodeAttributes, v101 => { nodeAttributes = v101; }).Val(true).Val("TrumpfAssetGeneral.CINumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.CINumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref10; }
            object hlContext_vref11 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref11, v102 => { hlContext_vref11 = v102; }).Ref(xmldoc, v103 => { xmldoc = v103; }).Ref(nodeAttributes, v104 => { nodeAttributes = v104; }).Val(false).Val("TrumpfAssetGeneral.CompanyCode").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.CompanyCode", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref11; }
            object hlContext_vref12 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref12, v105 => { hlContext_vref12 = v105; }).Ref(xmldoc, v106 => { xmldoc = v106; }).Ref(nodeAttributes, v107 => { nodeAttributes = v107; }).Val(false).Val("TrumpfAssetGeneral.InvestmentNumber").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.InvestmentNumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref12; }
            object hlContext_vref13 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref13, v108 => { hlContext_vref13 = v108; }).Ref(xmldoc, v109 => { xmldoc = v109; }).Ref(nodeAttributes, v110 => { nodeAttributes = v110; }).Val(false).Val("TrumpfAssetGeneral.Manufacturer").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.Manufacturer", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref13; }
            object hlContext_vref14 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref14, v111 => { hlContext_vref14 = v111; }).Ref(xmldoc, v112 => { xmldoc = v112; }).Ref(nodeAttributes, v113 => { nodeAttributes = v113; }).Val(false).Val("TrumpfAssetGeneral.SAPCostCenter").Val(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.SAPCostCenter", (Int16)0, (Int16)0, (Int16)0, (Int16)0)));
            }
            finally { hlContext = hlContext_vref14; }

            // Save to File
            if (_.IF(_.OR(_.EQ(_.NullableSTR(NewChangeObj), "0"), _.EQ(_.NullableSTR(NewChangeObj), ""))))
            {
                Filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\HELPLINE_out\\c11\\hlnew", objDefname, "_", _.CALLm1v0(this, _.NnO(hlObj, "hlObj"), "GetID"), ".xml");
            }
            else
            {
                Filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\HELPLINE_out\\c11\\hlchange", objDefname, "_", _.CALLm1v0(this, _.NnO(hlObj, "hlObj"), "GetID"), ".xml");
            }
            _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "Save", Filename);

        }
        public void ExportObjectIncident(ref object hlContext, ref object hlObj)
        {
            object objDefname = null;
            object aliasname1 = null;
            object aliasname2 = null;
            object aliasnameSU = null;
            object ElimierungsgrundDE = null;
            object ElimierungsgrundEN = null;
            object Buchungskreis = null;
            object Buchungskreis1 = null;
            object TeamKeyword = null;
            object Kontierungsnr = null;
            object Kontierungstyp = null;
            object Beschreibung = null;
            object cinummer = null;
            object increqsubject = null;
            object xmldoc = null;
            object nodeData = null;
            object nodeObjects = null;
            object nodeObject = null;
            object attAliasName = null;
            object nodeAttributes = null;
            object nodeServiceUnits = null;
            object nodeServiceUnit = null;
            object attAliasNameSU = null;
            object nodeObject2 = null;
            object attAliasName2 = null;
            object nodeAttributes2 = null;
            object nodeRelations = null;
            object nodeProduct2Case = null;
            object Filename = null;
            objDefname = "IncidentRequest";
            aliasname1 = "obj1";
            aliasname2 = "obj2";
            aliasnameSU = "objSU";
            ElimierungsgrundDE = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetStatus.CISubStatus", (Int16)7, (Int16)0, (Int16)0, (Int16)0));
            ElimierungsgrundEN = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetStatus.CISubStatus", (Int16)9, (Int16)0, (Int16)0, (Int16)0));
            Buchungskreis = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.CompanyCode", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Buchungskreis1 = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.CompanyCode", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            TeamKeyword = "";
            Kontierungsnr = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.AllocationNumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Kontierungstyp = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "ProcurementDetail.AllocationType", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            Beschreibung = "";
            Beschreibung = _.CONCAT("CI ist auf Status 'Elimiert' gesetzt worden. Die CI-Nummmer steht im Betreff. Der Eliminierungsgrund lautet: ", ElimierungsgrundDE);
            Beschreibung = _.CONCAT(Beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "The CI-Status is set to Eliminated. The CI-Number is displayed in the subject of the incident. The elimination reason is: ", ElimierungsgrundEN);
            Beschreibung = _.CONCAT(Beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Kontierungsnummer: ", Kontierungsnr);
            Beschreibung = _.CONCAT(Beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Kontierungstyp: ", Kontierungstyp);
            Beschreibung = _.CONCAT(Beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Allocationnumber: ", Kontierungsnr);
            Beschreibung = _.CONCAT(Beschreibung, _.CHR((Int16)13), _.CHR((Int16)10), "Allocationtype: ", Kontierungstyp);

            if (_.IF(_.EQ(Buchungskreis, "107")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "110")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "111")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "114")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "122")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "146")))
            {
                TeamKeyword = "KOControllingDitzingen";
            }
            else if (_.IF(_.EQ(Buchungskreis, "222")))
            {
                TeamKeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(Buchungskreis, "223")))
            {
                TeamKeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(Buchungskreis, "225")))
            {
                TeamKeyword = "KOControllingGruesch";
            }
            else if (_.IF(_.EQ(Buchungskreis, "314")))
            {
                TeamKeyword = "KOControllingPasching";
            }
            else if (_.IF(_.EQ(Buchungskreis, "231")))
            {
                TeamKeyword = "KOControllingFarmington";
            }
            else if (_.IF(_.EQ(Buchungskreis, "237")))
            {
                TeamKeyword = "KOControllingCranbury";
            }

            /////////////////////////////////////////////////////////////
            cinummer = _.VAL(_.CALLm1v5(this, _.NnO(hlObj, "hlObj"), "GetValue", "TrumpfAssetGeneral.CINumber", (Int16)0, (Int16)0, (Int16)0, (Int16)0));
            increqsubject = _.CONCAT("Eliminierung/Elimination: ", cinummer, " Internal helpLine-ID: ", _.CALLm1v0(this, _.NnO(hlObj, "hlObj"), "GetID"));

            // VBScript source code
            xmldoc = _.CREATEOBJECT("msxml2.DomDocument");

            //create root element
            nodeData = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Data")));
            nodeObjects = _.OBJ(_.CALLm1v1(this, _.NnO(nodeData, "nodeData"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Objects")));

            ////// obj1: IncidentRequest///////////////////////////////////////////////////////

            nodeObject = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObjects, "nodeObjects"), "appendChild", _.CALLm1argp(this, _.NnO(xmldoc, "xmldoc"), "createElement", _.ARGS.Ref(objDefname, v114 => { objDefname = v114; }))));
            attAliasName = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createAttribute", "aliasname"));
            _.SETm1a0(this, _.NnO(attAliasName, "attAliasName"), "Text", _.VAL(aliasname1));
            _.CALLm2argp(this, _.NnO(nodeObject, "nodeObject"), "Attributes", "setNamedItem", _.ARGS.Ref(attAliasName, v115 => { attAliasName = v115; }));
            nodeAttributes = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObject, "nodeObject"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Attributes")));
            nodeServiceUnits = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObject, "nodeObject"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "ServiceUnits")));
            nodeServiceUnit = _.OBJ(_.CALLm1v1(this, _.NnO(nodeServiceUnits, "nodeServiceUnits"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "ServiceUnit")));
            attAliasNameSU = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createAttribute", "aliasname"));
            _.SETm1a0(this, _.NnO(attAliasNameSU, "attAliasNameSU"), "Text", _.VAL(aliasnameSU));
            _.CALLm2argp(this, _.NnO(nodeServiceUnit, "nodeServiceUnit"), "Attributes", "setNamedItem", _.ARGS.Ref(attAliasNameSU, v116 => { attAliasNameSU = v116; }));
            ////////////////// HLOBJECT.ID
            //Call AppendNode(hlContext,xmldoc, nodeAttributes, True, "HLOBJECTINFO.ID", hlObj.GetValue("HLOBJECTINFO.ID", 0,0,0,0)) '
            object hlContext_vref15 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref15, v117 => { hlContext_vref15 = v117; }).Ref(xmldoc, v118 => { xmldoc = v118; }).Ref(nodeAttributes, v119 => { nodeAttributes = v119; }).Val(true).Val("CaseGeneral.Subject").Ref(increqsubject, v120 => { increqsubject = v120; }));
            }
            finally { hlContext = hlContext_vref15; }
            object hlContext_vref16 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref16, v121 => { hlContext_vref16 = v121; }).Ref(xmldoc, v122 => { xmldoc = v122; }).Ref(nodeAttributes, v123 => { nodeAttributes = v123; }).Val(false).Val("CaseDescription.DescriptionText").Ref(Beschreibung, v124 => { Beschreibung = v124; }));
            }
            finally { hlContext = hlContext_vref16; }
            object hlContext_vref17 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref17, v125 => { hlContext_vref17 = v125; }).Ref(xmldoc, v126 => { xmldoc = v126; }).Ref(nodeAttributes, v127 => { nodeAttributes = v127; }).Val(false).Val("Keywords.KeywordOrga").Ref(TeamKeyword, v128 => { TeamKeyword = v128; }));
            }
            finally { hlContext = hlContext_vref17; }
            object hlContext_vref18 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref18, v129 => { hlContext_vref18 = v129; }).Ref(xmldoc, v130 => { xmldoc = v130; }).Ref(nodeAttributes, v131 => { nodeAttributes = v131; }).Val(false).Val("CaseGeneral.CompanyCode").Ref(Buchungskreis1, v132 => { Buchungskreis1 = v132; }));
            }
            finally { hlContext = hlContext_vref18; }
            object hlContext_vref19 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref19, v133 => { hlContext_vref19 = v133; }).Ref(xmldoc, v134 => { xmldoc = v134; }).Ref(nodeAttributes, v135 => { nodeAttributes = v135; }).Val(false).Val("Keywords.Keyword").Val("KWStdSWhelplineInterfaceAM"));
            }
            finally { hlContext = hlContext_vref19; }
            object hlContext_vref20 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref20, v136 => { hlContext_vref20 = v136; }).Ref(xmldoc, v137 => { xmldoc = v137; }).Ref(nodeAttributes, v138 => { nodeAttributes = v138; }).Val(false).Val("IncidentAttribute.IncidentStatus").Val("IncidentStatusNew"));
            }
            finally { hlContext = hlContext_vref20; }
            object hlContext_vref21 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref21, v139 => { hlContext_vref21 = v139; }).Ref(xmldoc, v140 => { xmldoc = v140; }).Ref(nodeAttributes, v141 => { nodeAttributes = v141; }).Val(false).Val("IncidentAttribute.RequestType").Val("RequestTypeService"));
            }
            finally { hlContext = hlContext_vref21; }
            object hlContext_vref22 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref22, v142 => { hlContext_vref22 = v142; }).Ref(xmldoc, v143 => { xmldoc = v143; }).Ref(nodeServiceUnit, v144 => { nodeServiceUnit = v144; }).Val(true).Val("IncidentSUAttribute.IncidentOperation").Val("IncidentOperation"));
            }
            finally { hlContext = hlContext_vref22; }

            ////// obj2: Product///////////////////////////////////////////////////////
            nodeObject2 = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObjects, "nodeObjects"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", _.CALLm1argp(this, _.NnO(hlObj, "hlObj"), "GetType", _.ARGS.ForceBrackets()))));
            attAliasName2 = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createAttribute", "aliasname"));
            _.SETm1a0(this, _.NnO(attAliasName2, "attAliasName2"), "Text", _.VAL(aliasname2));
            _.CALLm2argp(this, _.NnO(nodeObject2, "nodeObject2"), "Attributes", "setNamedItem", _.ARGS.Ref(attAliasName2, v145 => { attAliasName2 = v145; }));
            nodeAttributes2 = _.OBJ(_.CALLm1v1(this, _.NnO(nodeObject2, "nodeObject2"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Attributes")));
            object hlContext_vref23 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendNode", _.ARGS.Ref(hlContext_vref23, v146 => { hlContext_vref23 = v146; }).Ref(xmldoc, v147 => { xmldoc = v147; }).Ref(nodeAttributes2, v148 => { nodeAttributes2 = v148; }).Val(true).Val("TrumpfAssetGeneral.CINumber").Ref(cinummer, v149 => { cinummer = v149; }));
            }
            finally { hlContext = hlContext_vref23; }

            ////// Relations///////////////////////////////////////////////////////
            nodeRelations = _.OBJ(_.CALLm1v1(this, _.NnO(nodeData, "nodeData"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Relations")));
            nodeProduct2Case = _.OBJ(_.CALLm1v1(this, _.NnO(nodeRelations, "nodeRelations"), "appendChild", _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createElement", "Product2Case")));

            object hlContext_vref24 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendTextNode", _.ARGS.Ref(hlContext_vref24, v150 => { hlContext_vref24 = v150; }).Ref(xmldoc, v151 => { xmldoc = v151; }).Ref(nodeProduct2Case, v152 => { nodeProduct2Case = v152; }).Val("Parent").Ref(aliasnameSU, v153 => { aliasnameSU = v153; }));
            }
            finally { hlContext = hlContext_vref24; }
            object hlContext_vref25 = hlContext;
            try
            {
                _.CALLm1argp(this, _outer, "AppendTextNode", _.ARGS.Ref(hlContext_vref25, v154 => { hlContext_vref25 = v154; }).Ref(xmldoc, v155 => { xmldoc = v155; }).Ref(nodeProduct2Case, v156 => { nodeProduct2Case = v156; }).Val("Child").Ref(aliasname2, v157 => { aliasname2 = v157; }));
            }
            finally { hlContext = hlContext_vref25; }

            // Save to File
            Filename = _.CONCAT("\\\\srvditz1\\pi_intern\\helpline\\helpline_in\\c11\\", objDefname, "_", _.CALLm1v0(this, _.NnO(hlObj, "hlObj"), "GetID"), ".xml");

            _.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "Save", Filename);

        }
        public void AppendNode(ref object hlContext, ref object xmldoc, ref object nodeObject, ref object iskey, ref object key, ref object value)
        {
            object valueNode = null;
            object cdata = null;
            object attIsKey = null;
            object key_vref = key;
            try
            {
                valueNode = _.OBJ(_.CALLm1argp(this, _.NnO(xmldoc, "xmldoc"), "createElement", _.ARGS.Ref(key_vref, v158 => { key_vref = v158; })));
            }
            finally { key = key_vref; }
            object value_vref = value;
            try
            {
                cdata = _.OBJ(_.CALLm1argp(this, _.NnO(xmldoc, "xmldoc"), "createCDATASection", _.ARGS.Ref(value_vref, v159 => { value_vref = v159; })));
            }
            finally { value = value_vref; }
            _.CALLm1v1(this, _.NnO(valueNode, "valueNode"), "appendChild", cdata);
            _.CALLm1v1(this, _.NnO(nodeObject, "nodeObject"), "appendChild", valueNode);

            attIsKey = _.OBJ(_.CALLm1v1(this, _.NnO(xmldoc, "xmldoc"), "createAttribute", "iskey"));
            if (_.IF(iskey))
            {
                _.SETm1a0(this, _.NnO(attIsKey, "attIsKey"), "Text", "true");
            }
            else
            {
                _.SETm1a0(this, _.NnO(attIsKey, "attIsKey"), "Text", "false");
            }
            _.CALLm2argp(this, _.NnO(valueNode, "valueNode"), "Attributes", "setNamedItem", _.ARGS.Ref(attIsKey, v160 => { attIsKey = v160; }));

        }
        public void AppendTextNode(ref object hlContext, ref object xmldoc, ref object nodeObject, ref object key, ref object value)
        {
            object valueNode = null;
            object key_vref2 = key;
            try
            {
                valueNode = _.OBJ(_.CALLm1argp(this, _.NnO(xmldoc, "xmldoc"), "createElement", _.ARGS.Ref(key_vref2, v161 => { key_vref2 = v161; })));
            }
            finally { key = key_vref2; }
            _.CALLm1v1(this, _.NnO(nodeObject, "nodeObject"), "appendChild", valueNode);

            _.SETm1a0(this, _.NnO(valueNode, "valueNode"), "Text", _.VAL(value));
        }
        public object DBConnectionString(ref object hlContext)
        {
            object DBConnectionString_retVal = null;
            const string DBConnection = "Provider=SQLOLEDB.1;Password=helplinedata;Persist Security Info=True;User ID=helplinedata;Initial Catalog=helpline_data;Data Source=srv01itsm2";

            return DBConnectionString_retVal;
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
