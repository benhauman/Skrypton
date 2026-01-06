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

            _.CALL(this, _outer, "ButtonGeneralInfo_Click");
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

        public object convertsize(ref object size)
        {
            object ConvertSize_retVal = null;
            object suffix = null; /* Undeclared in source */

            //MsgBox "Converting Size for " & Size
            size = _.CSNG(_.REPLACE(size, ",", ""));

            if (_.IF(_.NOT(_.EQ(_.VARTYPE(size), VBScriptConstants.vbSingle))))
            {
                ConvertSize_retVal = "SIZE INPUT ERROR";
                return ConvertSize_retVal;
            }

            suffix = " B";
            if (_.IF(_.GTE(_.NullableNUM(size), (Int16)1024)))
            {
                suffix = " KB";
            }
            if (_.IF(_.GTE(_.NullableNUM(size), 1048576)))
            {
                suffix = " MB";
            }
            if (_.IF(_.GTE(_.NullableNUM(size), 1073741824)))
            {
                suffix = " GB";
            }
            if (_.IF(_.GTE(_.NullableNUM(size), 1099511627776d)))
            {
                suffix = " TB";
            }

            //Select Case Suffix
            //	Case " KB" Size = Round(Size / 1024, 2)
            //	Case " MB" Size = Round(Size / 1048576, 2)
            //	Case " GB" Size = Round(Size / 1073741824, 2)
            //	Case " TB" Size = Round(Size / 1099511627776, 2)
            //End Select
            //
            ConvertSize_retVal = _.CONCAT(size, suffix);
            return ConvertSize_retVal;
        }

        public object getnexthinkuser()
        {
            return "myusr2";
        }

        public object getnexthinkbaseurl()
        {
            return "";
        }

        public object getnexthinkpassword()
        {
            return "mypwd2";
        }

        public void buttongeneralinfo_click()
        {
            var errOn = _.GETERRORTRAPPINGTOKEN();
            object nexthinkbaseurl = null;
            object nexthinkquery = null;
            object nexthinkurl = null;
            object colorwarning = null;
            object colorcheck = null;
            object hostname = null;
            object xmlhttp = null;
            object dict = null;
            object curnode = null;
            object i = null;
            object xmldoc = null; /* Undeclared in source */
            object n = null; /* Undeclared in source */

            _.SET((Int16)1, this, _env.tabpagegeneralinfo, "ShowControl");
            _.SET((Int16)3, this, _env.tabpagesoftwareoshealth, "ShowControl");
            _.SET((Int16)3, this, _env.tabpagesecuritycompliance, "ShowControl");
            _.SET((Int16)3, this, _env.tabpagetechnicalinfo, "ShowControl");
            _.SET((Int16)3, this, _env.tabpagenetworkhealth, "ShowControl");
            _.SET((Int16)3, this, _env.tabpagel1checklist, "ShowControl");

            _.SET((Int16)1, this, _env.groupboxgeneralinfo, "ShowControl");
            _.SET((Int16)3, this, _env.groupboxtechnicalinfo, "ShowControl");
            _.SET((Int16)3, this, _env.groupboxsoftwareoshealth, "ShowControl");
            _.SET((Int16)3, this, _env.groupboxsecuritycompliance, "ShowControl");
            _.SET((Int16)3, this, _env.groupboxnetworkhealth, "ShowControl");
            _.SET((Int16)3, this, _env.groupboxl1checklist, "ShowControl");

            _.SET("#5b5b5b", this, _env.buttongeneralinfo, "BackColor");
            _.SET("#1B709F", this, _env.buttontechnicalinfo, "BackColor");
            _.SET("#1B709F", this, _env.buttonswhealth, "BackColor");
            _.SET("#1B709F", this, _env.buttonsecuritycompliance, "BackColor");
            _.SET("#1B709F", this, _env.buttonnetworkhealth, "BackColor");
            _.SET("#1B709F", this, _env.buttonl1checklist, "BackColor");

            _.SET((Int16)1, this, _env.tabcontrolnexthink, "ShowControl");
            _.SET(true, this, _env.tabpagegeneralinfo, "RequestFocus");

            //Clear TextBoxes

            _.SET("", this, _env.textboxgeneralcalltime, "Text");
            _.SET("", this, _env.textboxgeneralhostname, "Text");
            _.SET("", this, _env.textboxlgeneraldevicemanufacturer, "Text");
            _.SET("", this, _env.textboxgeneraldeviceproductversion, "Text");
            _.SET("", this, _env.textboxgenerallastip, "Text");
            _.SET("", this, _env.textboxgeneralgroupname, "Text");
            _.SET("", this, _env.textboxgeneralos, "Text");
            _.SET("", this, _env.textboxgenerallastboottime, "Text");
            _.SET("", this, _env.textboxgenerallastlogon, "Text");
            _.SET("", this, _env.textboxgeneraldevicetype, "Text");
            _.SET("", this, _env.textboxgeneralbiosserialnumber, "Text");
            _.SET("", this, _env.textboxgeneralcpumodel, "Text");
            _.SET("", this, _env.textboxgeneralnumberofcpus, "Text");
            _.SET("", this, _env.textboxgeneralnumberoflogprocs, "Text");
            _.SET("", this, _env.textboxgeneralnumberofcores, "Text");
            _.SET("", this, _env.textboxgeneralcpufreq, "Text");
            _.SET("", this, _env.textboxgeneraltotalram, "Text");
            _.SET("", this, _env.textboxgeneralnumberofgraphcards, "Text");

            // --- GroupBoxTechnicalInfo

            _.SET("", this, _env.textboxtechnicalinfototaldrivecapnow, "Text");
            _.SET("", this, _env.textboxtechnicalinfototalfreespacenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfototaldriveusagenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfosystemdrivecapnow, "Text");
            _.SET("", this, _env.textboxtechnicalinfosystemdrivefreespacenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighcputimenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighmemorytimenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighiotimenow, "Text");
            _.SET("", this, _env.textboxtechnicalinfototaldrivecap7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfototalfreespace7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfototaldriveusage7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfosystemdrivecap7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfosystemdrivefreespace7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighcputime7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighmemorytime7days, "Text");
            _.SET("", this, _env.textboxtechnicalinfohighiotime7days, "Text");

            _.SET((Int16)3, this, _env.imagenoktechnicalinfototalfreespacenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfototalfreespacenow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfototaldriveusagenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfototaldriveusagenow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfosystemdrivecapnow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfosystemdrivecapnow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfosystemdrivefreespacenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfosystemdrivefreespacenow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighcputimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighcputimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighmemorytimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighmemorytimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighiotimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighiotimenow, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfototalfreespace7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfototalfreespace7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfototaldriveusage7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfototaldriveusage7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfosystemdrivecap7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfosystemdrivecap7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfosystemdrivefreespace7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfosystemdrivefreespace7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighcputime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighcputime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighmemorytime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighmemorytime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoktechnicalinfohighiotime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoktechnicalinfohighiotime7days, "ShowControl");

            // GroupBox Software OS Health

            _.SET("", this, _env.textboxsoftwareoshealthosversionarchitecture, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthosname, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthwmistatus, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthlastsystemupdate, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthwindowsupdatestatus, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthnumberofapps, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthnumberofexes, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthnumberofbins, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthosendofsupport, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthosie11support, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthwin10ready, "Text");
            _.SET("", this, _env.textboxsoftwareoshealthoscomplience, "Text");

            // GroupBox Security Compliance

            _.SET("", this, _env.textboxsecuritycomplianceinetsecuritysettings, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceuseraccountstatus, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantivirusname, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantivirusrtp, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantivirusuptodate, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantivirusnumber, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantivirusall, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantispywarename, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantispywarertp, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantispywareuptodate, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantispywarenumber, "Text");
            _.SET("", this, _env.textboxsecuritycomplianceantispywareall, "Text");
            _.SET("", this, _env.textboxsecuritycompliancefirewallname, "Text");
            _.SET("", this, _env.textboxsecuritycompliancefirewallrtp, "Text");
            _.SET("", this, _env.textboxsecuritycompliancefirewallnumber, "Text");
            _.SET("", this, _env.textboxsecuritycompliancefirewallall, "Text");

            // GroupBox Network Health
            _.SET("", this, _env.textboxnetworkhealthincomingnettaffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthoutgoingnettaffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthtotalnettaffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthsuccessnetconnectionratio24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthnetavaillevel24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgincomingnetbitrate24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgoutgoingnetbitrate24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgnetresponsetime24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthincomingwebtraffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthoutgoingwebtraffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthtotalwebtraffic24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgincomingwebbitrate24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgoutgoingwebbitrate24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgwebrequestsize24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgwebresponsesize24hours, "Text");
            _.SET("", this, _env.textboxnetworkhealthsuccesshttprequestratio24hours, "Text");

            _.SET("", this, _env.textboxnetworkhealthincomingnettaffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthoutgoingnettaffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthtotalnettaffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthsuccessnetconnectionratio7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthnetavaillevel7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgincomingnetbitrate7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgoutgoingnetbitrate7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgnetresponsetime7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthincomingwebtraffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthoutgoingwebtraffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthtotalwebtraffic7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgincomingwebbitrate7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgoutgoingwebbitrate7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgwebrequestsize7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthavgwebresponsesize7days, "Text");
            _.SET("", this, _env.textboxnetworkhealthsuccesshttprequestratio7days, "Text");

            _.SET((Int16)3, this, _env.imageoknetworkhealthincomingnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthincomingnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthincomingnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthincomingnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthoutgoingnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthoutgoingnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthoutgoingnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthoutgoingnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthtotalnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthtotalnettaffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthtotalnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthtotalnettaffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthsuccessnetconnectionratio24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthsuccessnetconnectionratio24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthsuccessnetconnectionratio7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthsuccessnetconnectionratio7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthnetavaillevel24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthnetavaillevel24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthnetavaillevel7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthnetavaillevel7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgincomingnetbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgincomingnetbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgincomingnetbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgincomingnetbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgoutgoingnetbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgoutgoingnetbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgoutgoingnetbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgoutgoingnetbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgnetresponsetime24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgnetresponsetime24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgnetresponsetime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgnetresponsetime7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthincomingwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthincomingwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthincomingwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthincomingwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthoutgoingwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthoutgoingwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthoutgoingwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthoutgoingwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthtotalwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthtotalwebtraffic24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthtotalwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthtotalwebtraffic7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgincomingwebbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgincomingwebbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgincomingwebbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgincomingwebbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgoutgoingwebbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgoutgoingwebbitrate24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgoutgoingwebbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgoutgoingwebbitrate7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgwebrequestsize24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgwebrequestsize24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgwebrequestsize7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgwebrequestsize7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgwebresponsesize24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgwebresponsesize24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthavgwebresponsesize7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthavgwebresponsesize7days, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthsuccesshttprequestratio24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthsuccesshttprequestratio24hours, "ShowControl");
            _.SET((Int16)3, this, _env.imageoknetworkhealthsuccesshttprequestratio7days, "ShowControl");
            _.SET((Int16)3, this, _env.imagenoknetworkhealthsuccesshttprequestratio7days, "ShowControl");

            // GroupBox L1-Checkliste

            _.SET("", this, _env.textboxl1freespace, "Text");
            _.SET("", this, _env.textboxl1osuptodate, "Text");
            _.SET("", this, _env.textboxl1browser, "Text");
            _.SET("", this, _env.textboxl1collaboration, "Text");
            _.SET("", this, _env.textboxl1antivirus, "Text");
            _.SET("", this, _env.textboxl1antivirus2, "Text");
            _.SET("", this, _env.textboxl1antivirus3, "Text");
            _.SET("", this, _env.textboxl1defender, "Text");
            _.SET("", this, _env.textboxl1bootlogon2, "Text");
            _.SET("", this, _env.textboxl1bootlogon3, "Text");
            _.SET("", this, _env.textboxl1cpu24, "Text");
            _.SET("", this, _env.textboxl1cpu7, "Text");
            _.SET("", this, _env.textboxl1speicher24, "Text");
            _.SET("", this, _env.textboxl1speicher7, "Text");
            _.SET("", this, _env.textboxl1bluescreen24, "Text");
            _.SET("", this, _env.textboxl1bluescrren7, "Text");
            _.SET("", this, _env.textboxl1hardreset24, "Text");
            _.SET("", this, _env.textboxl1hardreset7, "Text");

            // --- GroupBoxGeneralInfo

            nexthinkbaseurl = _.CONCAT(_.CALL(this, _outer, "getNexthinkBaseURL", _.ARGS.ForceBrackets()), "query?p1=");
            nexthinkquery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml";

            colorwarning = "#F20012";
            colorcheck = "#1B709F";

            hostname = _.VAL(_.CALL(this, _env.hlobj, "GetValue", _.ARGS.Val("ComputerDetail.Hostname").Val((Int16)0).Val((Int16)0).Val((Int16)0).Val((Int16)0)));

            if (_.IF(_.EQ(_.NullableSTR(hostname), "")))
            {
                _.CALL(this, _env.model, "MsgBox", _.ARGS.Val("Der Computer hat keinen Hostnamen."));
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            nexthinkurl = _.CONCAT(nexthinkbaseurl, _.UCASE(hostname), nexthinkquery);
            nexthinkurl = "https://httpbin.org/get";

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);

            //MsgBox nexthinkURL

            //time of call
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, _, "FORMATDATETIME", _.ARGS.Val(_.NOW()).Val(VBScriptConstants.vbGeneralDate))), this, _env.textboxgeneralcalltime, "Text");
            });

            _.HANDLEERROR(errOn, () => {
                xmlhttp = _.OBJ(_.CREATEOBJECT("Msxml2.ServerXMLHTTP.6.0"));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALL(this, xmlhttp, "setOption", _.ARGS.Val((Int16)2).Val((Int16)13056));
            }); //bypass certificate errors
            _.HANDLEERROR(errOn, () => {
                _.CALL(this, xmlhttp, "open", _.ARGS.Val("GET").Ref(nexthinkurl, v => { nexthinkurl = v; }).Val(false).Val(_.CALL(this, _outer, "getNexthinkUser", _.ARGS.ForceBrackets())).Val(_.CALL(this, _outer, "getNexthinkPassword", _.ARGS.ForceBrackets())));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALL(this, xmlhttp, "send");
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALL(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.model, "MsgBox", _.ARGS.Val("Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar."));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.model, "MsgBox", _.ARGS.Val(_.CONCAT("Error Description: ", _.CALL(this, _.ERR, "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALL(this, _.ERR, "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALL(this, _.ERR, "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALL(this, _.ERR, "HelpContext"))));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            //Reset the Error Data
            _.HANDLEERROR(errOn, () => {
                _.CLEARANYERROR();
            });

            _.HANDLEERROR(errOn, () => {
                xmldoc = _.OBJ(_.CREATEOBJECT("Msxml2.DOMDocument"));
            });
            _.HANDLEERROR(errOn, () => {
                _.SET("false", this, xmldoc, "async");
            });
            _.HANDLEERROR(errOn, () => {
                _.CALL(this, xmldoc, "load", _.ARGS.Val(_.CALL(this, xmlhttp, "responseXML")));
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALL(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.model, "MsgBox", _.ARGS.Val("Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten."));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.model, "MsgBox", _.ARGS.Val(_.CONCAT("Error Description: ", _.CALL(this, _.ERR, "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALL(this, _.ERR, "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALL(this, _.ERR, "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALL(this, _.ERR, "HelpContext"))));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            _.HANDLEERROR(errOn, () => {
                dict = _.OBJ(_.CREATEOBJECT("Scripting.Dictionary"));
            });

            //iterate all nodes and write into dictionary
            _.HANDLEERROR(errOn, () => {
                i = (Int16)0;
            });
            IEnumerator enumerationContent = null;
            _.HANDLEERROR(errOn, () => {
                enumerationContent = _.ENUMERABLE(_.CALL(this, xmldoc, "SelectNodes", _.ARGS.Val("//table/header/*"))).GetEnumerator();
            });
            while (true)
            {
                if (enumerationContent != null)
                {
                    if (!enumerationContent.MoveNext())
                        break;
                    n = enumerationContent.Current;
                }
                _.HANDLEERROR(errOn, () => {
                    curnode = _.OBJ(_.CALL(this, xmldoc, "documentElement", "selectSingleNode", _.ARGS.Val(_.CONCAT("//table/body/r/c", i))));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, dict, "Add", _.ARGS.Val(_.CALL(this, n, "Text")).Val(_.CALL(this, curnode, "Text")));
                });
                _.HANDLEERROR(errOn, () => {
                    i = _.ADD(i, (Int16)1);
                });
                if (enumerationContent == null)
                    break;
            }

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALL(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALL(this, _env.model, "MsgBox", _.ARGS.Val("Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten."));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            // from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

            //fill textboxes
            //LabelNName.Text = dict.key("name")
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("name"))), this, _env.textboxgeneralhostname, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("last_ip_address"))), this, _env.textboxgenerallastip, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("device_manufacturer"))), this, _env.textboxlgeneraldevicemanufacturer, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("device_model"))), this, _env.textboxgeneraldeviceproductversion, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("os_version_and_architecture"))), this, _env.textboxgeneralos, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("group_name"))), this, _env.textboxgeneralgroupname, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, _, "FORMATDATETIME", _.ARGS.Val(_.REPLACE(_.CALL(this, dict, "Item", _.ARGS.Val("last_system_boot")), "T", " ")).Val(VBScriptConstants.vbGeneralDate))), this, _env.textboxgenerallastboottime, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, _, "FORMATDATETIME", _.ARGS.Val(_.REPLACE(_.CALL(this, dict, "Item", _.ARGS.Val("last_logon_time")), "T", " ")).Val(VBScriptConstants.vbGeneralDate))), this, _env.textboxgenerallastlogon, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("last_logged_on_user"))), this, _env.textboxgeneraldevicetype, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("bios_serial_number"))), this, _env.textboxgeneralbiosserialnumber, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("cpu_model"))), this, _env.textboxgeneralcpumodel, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("number_of_cpus"))), this, _env.textboxgeneralnumberofcpus, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("logical_cpu_number"))), this, _env.textboxgeneralnumberoflogprocs, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("number_of_cores"))), this, _env.textboxgeneralnumberofcores, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.CONCAT(_.CALL(this, dict, "Item", _.ARGS.Val("cpu_frequency")), " MHz"), this, _env.textboxgeneralcpufreq, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, _outer, "ConvertSize", _.ARGS.Val(_.CALL(this, dict, "Item", _.ARGS.Val("total_ram"))))), this, _env.textboxgeneraltotalram, "Text");
            });

            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, dict, "Item", _.ARGS.Val("number_of_graphical_cards"))), this, _env.textboxgeneralnumberofgraphcards, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALL(this, _outer, "ConvertSize", _.ARGS.Val(_.CALL(this, dict, "Item", _.ARGS.Val("graphical_card_ram"))))), this, _env.textboxgeneralgraphcardram, "Text");
            });

            _.RELEASEERRORTRAPPINGTOKEN(errOn);
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object buttongeneralinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object buttonl1checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object buttonnetworkhealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object buttonsecuritycompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object buttonswhealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object buttontechnicalinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxgeneralinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxl1checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxnetworkhealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxsecuritycompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxsoftwareoshealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object groupboxtechnicalinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlobj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgincomingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgincomingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgincomingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgincomingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgnetresponsetime24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgnetresponsetime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgoutgoingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgoutgoingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgoutgoingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgoutgoingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgwebrequestsize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgwebrequestsize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgwebresponsesize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthavgwebresponsesize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthincomingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthincomingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthincomingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthincomingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthnetavaillevel24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthnetavaillevel7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthoutgoingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthoutgoingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthoutgoingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthoutgoingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthsuccesshttprequestratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthsuccesshttprequestratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthsuccessnetconnectionratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthsuccessnetconnectionratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthtotalnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthtotalnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoknetworkhealthtotalwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighcputime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighcputimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighiotime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighiotimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighmemorytime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfohighmemorytimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfosystemdrivecap7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfosystemdrivecapnow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfosystemdrivefreespace7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfosystemdrivefreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfototaldriveusage7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfototaldriveusagenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imagenoktechnicalinfototalfreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgincomingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgincomingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgincomingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgincomingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgnetresponsetime24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgnetresponsetime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgoutgoingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgoutgoingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgoutgoingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgoutgoingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgwebrequestsize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgwebrequestsize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgwebresponsesize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthavgwebresponsesize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthincomingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthincomingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthincomingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthincomingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthnetavaillevel24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthnetavaillevel7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthoutgoingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthoutgoingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthoutgoingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthoutgoingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthsuccesshttprequestratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthsuccesshttprequestratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthsuccessnetconnectionratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthsuccessnetconnectionratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthtotalnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthtotalnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthtotalwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoknetworkhealthtotalwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighcputime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighcputimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighiotime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighiotimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighmemorytime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfohighmemorytimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfosystemdrivecap7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfosystemdrivecapnow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfosystemdrivefreespace7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfosystemdrivefreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfototaldriveusage7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfototaldriveusagenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfototalfreespace7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object imageoktechnicalinfototalfreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabcontrolnexthink { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagegeneralinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagel1checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagenetworkhealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagesecuritycompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagesoftwareoshealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object tabpagetechnicalinfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxchecklist2url { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralbiosserialnumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralcalltime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralcpufreq { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralcpumodel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneraldeviceproductversion { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneraldevicetype { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralgraphcardram { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralgroupname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralhostname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgenerallastboottime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgenerallastip { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgenerallastlogon { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralnumberofcores { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralnumberofcpus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralnumberofgraphcards { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralnumberoflogprocs { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneralos { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxgeneraltotalram { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1antivirus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1antivirus2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1antivirus3 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1bluescreen24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1bluescrren7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1bootlogon2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1bootlogon3 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1browser { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1collaboration { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1cpu24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1cpu7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1defender { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1freespace { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1hardreset24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1hardreset7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1osuptodate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1speicher24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxl1speicher7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxlgeneraldevicemanufacturer { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgincomingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgincomingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgincomingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgincomingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgnetresponsetime24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgnetresponsetime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgoutgoingnetbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgoutgoingnetbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgoutgoingwebbitrate24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgoutgoingwebbitrate7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgwebrequestsize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgwebrequestsize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgwebresponsesize24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthavgwebresponsesize7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthincomingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthincomingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthincomingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthincomingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthnetavaillevel24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthnetavaillevel7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthoutgoingnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthoutgoingnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthoutgoingwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthoutgoingwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthsuccesshttprequestratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthsuccesshttprequestratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthsuccessnetconnectionratio24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthsuccessnetconnectionratio7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthtotalnettaffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthtotalnettaffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthtotalwebtraffic24hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxnetworkhealthtotalwebtraffic7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantispywareall { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantispywarename { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantispywarenumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantispywarertp { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantispywareuptodate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantivirusall { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantivirusname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantivirusnumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantivirusrtp { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceantivirusuptodate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycompliancefirewallall { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycompliancefirewallname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycompliancefirewallnumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycompliancefirewallrtp { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceinetsecuritysettings { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsecuritycomplianceuseraccountstatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthlastsystemupdate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthnumberofapps { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthnumberofbins { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthnumberofexes { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthoscomplience { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthosendofsupport { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthosie11support { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthosname { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthosversionarchitecture { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthwin10ready { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthwindowsupdatestatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxsoftwareoshealthwmistatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighcputime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighcputimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighiotime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighiotimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighmemorytime7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfohighmemorytimenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfosystemdrivecap7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfosystemdrivecapnow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfosystemdrivefreespace7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfosystemdrivefreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototaldrivecap7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototaldrivecapnow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototaldriveusage7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototaldriveusagenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototalfreespace7days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object textboxtechnicalinfototalfreespacenow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}