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

            _.CALLm1v0(this, _outer, "ButtonGeneralInfo_Click");
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

        public object ConvertSize(ref object Size)
        {
            object ConvertSize_retVal = null;
            object Suffix = null; /* Undeclared in source */

            //MsgBox "Converting Size for " & Size
            Size = _.CSNG(_.REPLACE(Size, ",", ""));

            if (_.IF(_.NOT(_.EQ(_.VARTYPE(Size), VBScriptConstants.vbSingle))))
            {
                ConvertSize_retVal = "SIZE INPUT ERROR";
                return ConvertSize_retVal;
            }

            Suffix = " B";
            if (_.IF(_.GTE(_.NullableNUM(Size), (Int16)1024)))
            {
                Suffix = " KB";
            }
            if (_.IF(_.GTE(_.NullableNUM(Size), 1048576)))
            {
                Suffix = " MB";
            }
            if (_.IF(_.GTE(_.NullableNUM(Size), 1073741824)))
            {
                Suffix = " GB";
            }
            if (_.IF(_.GTE(_.NullableNUM(Size), 1099511627776d)))
            {
                Suffix = " TB";
            }

            if (_.IF(_.EQ(Suffix, " KB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _, "ROUND", _.DIV(Size, (Int16)1024), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " MB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _, "ROUND", _.DIV(Size, 1048576), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " GB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _, "ROUND", _.DIV(Size, 1073741824), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " TB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _, "ROUND", _.DIV(Size, 1099511627776d), (Int16)2));
            }

            ConvertSize_retVal = _.CONCAT(Size, Suffix);
            return ConvertSize_retVal;
        }

        public object getNexthinkUser()
        {
            return "myusr2";
        }

        public object getNexthinkBaseURL()
        {
            return "";
        }

        public object getNexthinkPassword()
        {
            return "mypwd2";
        }

        public void ButtonGeneralInfo_Click()
        {
            var errOn = _.GETERRORTRAPPINGTOKEN();
            object nexthinkBaseURL = null;
            object nexthinkQuery = null;
            object nexthinkURL = null;
            object colorWarning = null;
            object colorCheck = null;
            object hostname = null;
            object xmlhttp = null;
            object dict = null;
            object curnode = null;
            object i = null;
            object xmlDoc = null; /* Undeclared in source */
            object n = null; /* Undeclared in source */

            _.SET((Int16)1, this, _env.TabPageGeneralInfo, "ShowControl");
            _.SET((Int16)3, this, _env.TabPageSoftwareOSHealth, "ShowControl");
            _.SET((Int16)3, this, _env.TabPageSecurityCompliance, "ShowControl");
            _.SET((Int16)3, this, _env.TabPageTechnicalInfo, "ShowControl");
            _.SET((Int16)3, this, _env.TabPageNetworkHealth, "ShowControl");
            _.SET((Int16)3, this, _env.TabPageL1Checklist, "ShowControl");

            _.SET((Int16)1, this, _env.GroupBoxGeneralInfo, "ShowControl");
            _.SET((Int16)3, this, _env.GroupBoxTechnicalInfo, "ShowControl");
            _.SET((Int16)3, this, _env.GroupBoxSoftwareOSHealth, "ShowControl");
            _.SET((Int16)3, this, _env.GroupBoxSecurityCompliance, "ShowControl");
            _.SET((Int16)3, this, _env.GroupBoxNetworkHealth, "ShowControl");
            _.SET((Int16)3, this, _env.GroupBoxL1Checklist, "ShowControl");

            _.SET("#5b5b5b", this, _env.ButtonGeneralInfo, "BackColor");
            _.SET("#1B709F", this, _env.ButtonTechnicalInfo, "BackColor");
            _.SET("#1B709F", this, _env.ButtonSWHealth, "BackColor");
            _.SET("#1B709F", this, _env.ButtonSecurityCompliance, "BackColor");
            _.SET("#1B709F", this, _env.ButtonNetworkHealth, "BackColor");
            _.SET("#1B709F", this, _env.ButtonL1Checklist, "BackColor");

            _.SET((Int16)1, this, _env.TabControlNexthink, "ShowControl");
            _.SET(true, this, _env.TabPageGeneralInfo, "RequestFocus");

            //Clear TextBoxes

            _.SET("", this, _env.TextBoxGeneralCallTime, "Text");
            _.SET("", this, _env.TextBoxGeneralHostName, "Text");
            _.SET("", this, _env.TextBoxlGeneralDeviceManufacturer, "Text");
            _.SET("", this, _env.TextBoxGeneralDeviceProductVersion, "Text");
            _.SET("", this, _env.TextBoxGeneralLastIP, "Text");
            _.SET("", this, _env.TextBoxGeneralGroupName, "Text");
            _.SET("", this, _env.TextBoxGeneralOS, "Text");
            _.SET("", this, _env.TextBoxGeneralLastBootTime, "Text");
            _.SET("", this, _env.TextBoxGeneralLastLogon, "Text");
            _.SET("", this, _env.TextBoxGeneralDeviceType, "Text");
            _.SET("", this, _env.TextBoxGeneralBIOSSerialNumber, "Text");
            _.SET("", this, _env.TextBoxGeneralCPUModel, "Text");
            _.SET("", this, _env.TextBoxGeneralNumberOfCPUs, "Text");
            _.SET("", this, _env.TextBoxGeneralNumberOfLogProcs, "Text");
            _.SET("", this, _env.TextBoxGeneralNumberOfCores, "Text");
            _.SET("", this, _env.TextBoxGeneralCPUFreq, "Text");
            _.SET("", this, _env.TextBoxGeneralTotalRAM, "Text");
            _.SET("", this, _env.TextBoxGeneralNumberOfGraphCards, "Text");

            // --- GroupBoxTechnicalInfo

            _.SET("", this, _env.TextBoxTechnicalInfoTotalDriveCapNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoTotalFreeSpaceNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoTotalDriveUsageNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoSystemDriveCapNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoSystemDriveFreeSpaceNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighCPUTimeNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighMemoryTimeNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighIOTimeNow, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoTotalDriveCap7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoTotalFreeSpace7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoTotalDriveUsage7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoSystemDriveCap7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoSystemDriveFreeSpace7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighCPUTime7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighMemoryTime7Days, "Text");
            _.SET("", this, _env.TextBoxTechnicalInfoHighIOTime7Days, "Text");

            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoTotalFreeSpaceNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoTotalFreeSpaceNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoTotalDriveUsageNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoTotalDriveUsageNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoSystemDriveCapNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoSystemDriveCapNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoSystemDriveFreeSpaceNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoSystemDriveFreeSpaceNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighCPUTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighCPUTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighMemoryTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighMemoryTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighIOTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighIOTimeNow, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoTotalFreeSpace7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoTotalFreeSpace7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoTotalDriveUsage7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoTotalDriveUsage7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoSystemDriveCap7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoSystemDriveCap7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoSystemDriveFreeSpace7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoSystemDriveFreeSpace7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighCPUTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighCPUTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighMemoryTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighMemoryTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKTechnicalInfoHighIOTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKTechnicalInfoHighIOTime7Days, "ShowControl");

            // GroupBox Software OS Health

            _.SET("", this, _env.TextBoxSoftwareOSHealthOSVersionArchitecture, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthOSName, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthWMIStatus, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthLastSystemUpdate, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthWindowsUpdateStatus, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthNumberOfApps, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthNumberOfExes, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthNumberOfBins, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthOSEndOfSupport, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthOSIE11Support, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthWin10Ready, "Text");
            _.SET("", this, _env.TextBoxSoftwareOSHealthOSComplience, "Text");

            // GroupBox Security Compliance

            _.SET("", this, _env.TextBoxSecurityComplianceInetSecuritySettings, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceUserAccountStatus, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntivirusName, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntivirusRTP, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntivirusUpToDate, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntivirusNumber, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntivirusAll, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntispywareName, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntispywareRTP, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntispywareUpToDate, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntispywareNumber, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceAntispywareAll, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceFirewallName, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceFirewallRTP, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceFirewallNumber, "Text");
            _.SET("", this, _env.TextBoxSecurityComplianceFirewallAll, "Text");

            // GroupBox Network Health
            _.SET("", this, _env.TextBoxNetworkHealthIncomingNetTaffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthOutgoingNetTaffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthTotalNetTaffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthSuccessNetConnectionRatio24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthNetAvailLevel24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgIncomingNetBitrate24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgNetResponseTime24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthIncomingWebTraffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthOutgoingWebTraffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthTotalWebTraffic24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgIncomingWebBitrate24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgWebRequestSize24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgWebResponseSize24Hours, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours, "Text");

            _.SET("", this, _env.TextBoxNetworkHealthIncomingNetTaffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthOutgoingNetTaffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthTotalNetTaffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthSuccessNetConnectionRatio7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthNetAvailLevel7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgIncomingNetBitrate7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgOutgoingNetBitrate7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgNetResponseTime7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthIncomingWebTraffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthOutgoingWebTraffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthTotalWebTraffic7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgIncomingWebBitrate7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgOutgoingWebBitrate7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgWebRequestSize7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthAvgWebResponseSize7Days, "Text");
            _.SET("", this, _env.TextBoxNetworkHealthSuccessHTTPRequestRatio7Days, "Text");

            _.SET((Int16)3, this, _env.ImageOKNetworkHealthIncomingNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthIncomingNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthIncomingNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthIncomingNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthOutgoingNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthOutgoingNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthOutgoingNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthOutgoingNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthTotalNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthTotalNetTaffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthTotalNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthTotalNetTaffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthSuccessNetConnectionRatio24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthSuccessNetConnectionRatio7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthSuccessNetConnectionRatio7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthNetAvailLevel24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthNetAvailLevel24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthNetAvailLevel7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthNetAvailLevel7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgIncomingNetBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgIncomingNetBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgIncomingNetBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgOutgoingNetBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgNetResponseTime24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgNetResponseTime24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgNetResponseTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgNetResponseTime7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthIncomingWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthIncomingWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthIncomingWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthIncomingWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthOutgoingWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthOutgoingWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthOutgoingWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthOutgoingWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthTotalWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthTotalWebTraffic24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthTotalWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthTotalWebTraffic7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgIncomingWebBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgIncomingWebBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgIncomingWebBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgOutgoingWebBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgWebRequestSize24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgWebRequestSize24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgWebRequestSize7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgWebRequestSize7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgWebResponseSize24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgWebResponseSize24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthAvgWebResponseSize7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthAvgWebResponseSize7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours, "ShowControl");
            _.SET((Int16)3, this, _env.ImageOKNetworkHealthSuccessHTTPRequestRatio7Days, "ShowControl");
            _.SET((Int16)3, this, _env.ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days, "ShowControl");

            // GroupBox L1-Checkliste

            _.SET("", this, _env.TextBoxL1FreeSpace, "Text");
            _.SET("", this, _env.TextBoxL1OSUpToDate, "Text");
            _.SET("", this, _env.TextBoxL1Browser, "Text");
            _.SET("", this, _env.TextBoxL1Collaboration, "Text");
            _.SET("", this, _env.TextBoxL1Antivirus, "Text");
            _.SET("", this, _env.TextBoxL1Antivirus2, "Text");
            _.SET("", this, _env.TextBoxL1Antivirus3, "Text");
            _.SET("", this, _env.TextBoxL1Defender, "Text");
            _.SET("", this, _env.TextBoxL1BootLogon2, "Text");
            _.SET("", this, _env.TextBoxL1BootLogon3, "Text");
            _.SET("", this, _env.TextBoxL1CPU24, "Text");
            _.SET("", this, _env.TextBoxL1CPU7, "Text");
            _.SET("", this, _env.TextBoxL1Speicher24, "Text");
            _.SET("", this, _env.TextBoxL1Speicher7, "Text");
            _.SET("", this, _env.TextBoxL1Bluescreen24, "Text");
            _.SET("", this, _env.TextBoxL1Bluescrren7, "Text");
            _.SET("", this, _env.TextBoxL1HardReset24, "Text");
            _.SET("", this, _env.TextBoxL1HardReset7, "Text");

            // --- GroupBoxGeneralInfo

            nexthinkBaseURL = _.CONCAT(_.CALLm1argp(this, _outer, "getNexthinkBaseURL", _.ARGS.ForceBrackets()), "query?p1=");
            nexthinkQuery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml";

            colorWarning = "#F20012";
            colorCheck = "#1B709F";

            hostname = _.VAL(_.CALLm1v5(this, _env.hlobj, "GetValue", "ComputerDetail.Hostname", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            if (_.IF(_.EQ(_.NullableSTR(hostname), "")))
            {
                _.CALLm1v1(this, _env.model, "MsgBox", "Der Computer hat keinen Hostnamen.");
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            nexthinkURL = _.CONCAT(nexthinkBaseURL, _.UCASE(hostname), nexthinkQuery);
            nexthinkURL = "https://httpbin.org/get";

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);

            //MsgBox nexthinkURL

            //time of call
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v2(this, _, "FORMATDATETIME", _.NOW(), VBScriptConstants.vbGeneralDate)), this, _env.TextBoxGeneralCallTime, "Text");
            });

            _.HANDLEERROR(errOn, () => {
                xmlhttp = _.OBJ(_.CREATEOBJECT("Msxml2.ServerXMLHTTP.6.0"));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v2(this, xmlhttp, "setOption", (Int16)2, (Int16)13056);
            }); //bypass certificate errors
            _.HANDLEERROR(errOn, () => {
                _.CALLm1argp(this, xmlhttp, "open", _.ARGS.Val("GET").Ref(nexthinkURL, v => { nexthinkURL = v; }).Val(false).Val(_.CALLm1argp(this, _outer, "getNexthinkUser", _.ARGS.ForceBrackets())).Val(_.CALLm1argp(this, _outer, "getNexthinkPassword", _.ARGS.ForceBrackets())));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v0(this, xmlhttp, "send");
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model, "MsgBox", "Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model, "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.ERR, "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.ERR, "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.ERR, "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.ERR, "HelpContext")));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            //Reset the Error Data
            _.HANDLEERROR(errOn, () => {
                _.CLEARANYERROR();
            });

            _.HANDLEERROR(errOn, () => {
                xmlDoc = _.OBJ(_.CREATEOBJECT("Msxml2.DOMDocument"));
            });
            _.HANDLEERROR(errOn, () => {
                _.SET("false", this, xmlDoc, "async");
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, xmlDoc, "load", _.CALLm1v0(this, xmlhttp, "responseXML"));
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model, "MsgBox", "Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model, "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.ERR, "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.ERR, "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.ERR, "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.ERR, "HelpContext")));
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
                enumerationContent = _.ENUMERABLE(_.CALLm1v1(this, xmlDoc, "SelectNodes", "//table/header/*")).GetEnumerator();
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
                    curnode = _.OBJ(_.CALLm2v1(this, xmlDoc, "documentElement", "selectSingleNode", _.CONCAT("//table/body/r/c", i)));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v2(this, dict, "Add", _.CALLm1v0(this, n, "Text"), _.CALLm1v0(this, curnode, "Text"));
                });
                _.HANDLEERROR(errOn, () => {
                    i = _.ADD(i, (Int16)1);
                });
                if (enumerationContent == null)
                    break;
            }

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR, "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model, "MsgBox", "Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten.");
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            // from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

            //fill textboxes
            //LabelNName.Text = dict.key("name")
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "name")), this, _env.TextBoxGeneralHostName, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "last_ip_address")), this, _env.TextBoxGeneralLastIP, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "device_manufacturer")), this, _env.TextBoxlGeneralDeviceManufacturer, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "device_model")), this, _env.TextBoxGeneralDeviceProductVersion, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "os_version_and_architecture")), this, _env.TextBoxGeneralOS, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "group_name")), this, _env.TextBoxGeneralGroupName, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v2(this, _, "FORMATDATETIME", _.REPLACE(_.CALLm1v1(this, dict, "Item", "last_system_boot"), "T", " "), VBScriptConstants.vbGeneralDate)), this, _env.TextBoxGeneralLastBootTime, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v2(this, _, "FORMATDATETIME", _.REPLACE(_.CALLm1v1(this, dict, "Item", "last_logon_time"), "T", " "), VBScriptConstants.vbGeneralDate)), this, _env.TextBoxGeneralLastLogon, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "last_logged_on_user")), this, _env.TextBoxGeneralDeviceType, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "bios_serial_number")), this, _env.TextBoxGeneralBIOSSerialNumber, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "cpu_model")), this, _env.TextBoxGeneralCPUModel, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "number_of_cpus")), this, _env.TextBoxGeneralNumberOfCPUs, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "logical_cpu_number")), this, _env.TextBoxGeneralNumberOfLogProcs, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "number_of_cores")), this, _env.TextBoxGeneralNumberOfCores, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.CONCAT(_.CALLm1v1(this, dict, "Item", "cpu_frequency"), " MHz"), this, _env.TextBoxGeneralCPUFreq, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, dict, "Item", "total_ram"))), this, _env.TextBoxGeneralTotalRAM, "Text");
            });

            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, dict, "Item", "number_of_graphical_cards")), this, _env.TextBoxGeneralNumberOfGraphCards, "Text");
            });
            _.HANDLEERROR(errOn, () => {
                _.SET(_.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, dict, "Item", "graphical_card_ram"))), this, _env.TextBoxGeneralGraphCardRAM, "Text");
            });

            _.RELEASEERRORTRAPPINGTOKEN(errOn);
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object ButtonGeneralInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonL1Checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonNetworkHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSecurityCompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonSWHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ButtonTechnicalInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxGeneralInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxL1Checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxNetworkHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxSecurityCompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxSoftwareOSHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object GroupBoxTechnicalInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object hlobj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgIncomingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgIncomingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgNetResponseTime24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgNetResponseTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgWebRequestSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgWebRequestSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgWebResponseSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthAvgWebResponseSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthIncomingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthIncomingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthIncomingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthIncomingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthNetAvailLevel24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthNetAvailLevel7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthOutgoingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthOutgoingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthOutgoingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthOutgoingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthSuccessNetConnectionRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthTotalNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthTotalNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKNetworkHealthTotalWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighCPUTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighCPUTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighIOTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighIOTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighMemoryTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoHighMemoryTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoSystemDriveCap7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoSystemDriveCapNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoSystemDriveFreeSpace7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoSystemDriveFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoTotalDriveUsage7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoTotalDriveUsageNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageNOKTechnicalInfoTotalFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgIncomingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgIncomingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgIncomingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgIncomingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgNetResponseTime24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgNetResponseTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgOutgoingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgOutgoingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgWebRequestSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgWebRequestSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgWebResponseSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthAvgWebResponseSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthIncomingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthIncomingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthIncomingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthIncomingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthNetAvailLevel24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthNetAvailLevel7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthOutgoingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthOutgoingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthOutgoingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthOutgoingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthSuccessHTTPRequestRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthSuccessNetConnectionRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthSuccessNetConnectionRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthTotalNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthTotalNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthTotalWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKNetworkHealthTotalWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighCPUTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighCPUTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighIOTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighIOTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighMemoryTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoHighMemoryTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoSystemDriveCap7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoSystemDriveCapNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoSystemDriveFreeSpace7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoSystemDriveFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoTotalDriveUsage7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoTotalDriveUsageNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoTotalFreeSpace7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object ImageOKTechnicalInfoTotalFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object model { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabControlNexthink { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageGeneralInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageL1Checklist { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageNetworkHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageSecurityCompliance { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageSoftwareOSHealth { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TabPageTechnicalInfo { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxChecklist2URL { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralBIOSSerialNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralCallTime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralCPUFreq { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralCPUModel { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralDeviceProductVersion { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralDeviceType { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralGraphCardRAM { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralGroupName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralHostName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralLastBootTime { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralLastIP { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralLastLogon { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralNumberOfCores { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralNumberOfCPUs { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralNumberOfGraphCards { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralNumberOfLogProcs { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralOS { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxGeneralTotalRAM { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Antivirus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Antivirus2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Antivirus3 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Bluescreen24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Bluescrren7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1BootLogon2 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1BootLogon3 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Browser { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Collaboration { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1CPU24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1CPU7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Defender { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1FreeSpace { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1HardReset24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1HardReset7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1OSUpToDate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Speicher24 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxL1Speicher7 { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxlGeneralDeviceManufacturer { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgIncomingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgIncomingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgIncomingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgIncomingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgNetResponseTime24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgNetResponseTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgOutgoingNetBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgOutgoingWebBitrate7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgWebRequestSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgWebRequestSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgWebResponseSize24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthAvgWebResponseSize7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthIncomingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthIncomingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthIncomingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthIncomingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthNetAvailLevel24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthNetAvailLevel7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthOutgoingNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthOutgoingNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthOutgoingWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthOutgoingWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthSuccessHTTPRequestRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthSuccessNetConnectionRatio24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthSuccessNetConnectionRatio7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthTotalNetTaffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthTotalNetTaffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthTotalWebTraffic24Hours { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxNetworkHealthTotalWebTraffic7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntispywareAll { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntispywareName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntispywareNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntispywareRTP { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntispywareUpToDate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntivirusAll { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntivirusName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntivirusNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntivirusRTP { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceAntivirusUpToDate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceFirewallAll { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceFirewallName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceFirewallNumber { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceFirewallRTP { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceInetSecuritySettings { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSecurityComplianceUserAccountStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthLastSystemUpdate { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthNumberOfApps { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthNumberOfBins { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthNumberOfExes { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthOSComplience { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthOSEndOfSupport { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthOSIE11Support { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthOSName { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthOSVersionArchitecture { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthWin10Ready { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthWindowsUpdateStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxSoftwareOSHealthWMIStatus { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighCPUTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighCPUTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighIOTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighIOTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighMemoryTime7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoHighMemoryTimeNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoSystemDriveCap7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoSystemDriveCapNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoSystemDriveFreeSpace7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoSystemDriveFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalDriveCap7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalDriveCapNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalDriveUsage7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalDriveUsageNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalFreeSpace7Days { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
        public object TextBoxTechnicalInfoTotalFreeSpaceNow { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
