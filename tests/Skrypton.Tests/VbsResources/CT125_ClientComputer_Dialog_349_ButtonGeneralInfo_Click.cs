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
                Size = _.ROUND(_.DIV(Size, (Int16)1024), 2);
            }
            else if (_.IF(_.EQ(Suffix, " MB")))
            {
                Size = _.ROUND(_.DIV(Size, 1048576), 2);
            }
            else if (_.IF(_.EQ(Suffix, " GB")))
            {
                Size = _.ROUND(_.DIV(Size, 1073741824), 2);
            }
            else if (_.IF(_.EQ(Suffix, " TB")))
            {
                Size = _.ROUND(_.DIV(Size, 1099511627776d), 2);
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
            int errOn = _.GETERRORTRAPPINGTOKEN();
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

            _.SETm1a0(this, _.NnO(_env.TabPageGeneralInfo, null), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.TabPageSoftwareOSHealth, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageSecurityCompliance, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageTechnicalInfo, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageNetworkHealth, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageL1Checklist, null), "ShowControl", (Int16)3);

            _.SETm1a0(this, _.NnO(_env.GroupBoxGeneralInfo, null), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.GroupBoxTechnicalInfo, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxSoftwareOSHealth, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxSecurityCompliance, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxNetworkHealth, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxL1Checklist, null), "ShowControl", (Int16)3);

            _.SETm1a0(this, _.NnO(_env.ButtonGeneralInfo, null), "BackColor", "#5b5b5b");
            _.SETm1a0(this, _.NnO(_env.ButtonTechnicalInfo, null), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonSWHealth, null), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonSecurityCompliance, null), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonNetworkHealth, null), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonL1Checklist, null), "BackColor", "#1B709F");

            _.SETm1a0(this, _.NnO(_env.TabControlNexthink, null), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.TabPageGeneralInfo, null), "RequestFocus", true);

            //Clear TextBoxes

            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCallTime, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralHostName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxlGeneralDeviceManufacturer, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceProductVersion, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastIP, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGroupName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralOS, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastBootTime, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastLogon, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceType, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralBIOSSerialNumber, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUModel, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCPUs, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfLogProcs, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCores, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUFreq, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralTotalRAM, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfGraphCards, null), "Text", "");

            // --- GroupBoxTechnicalInfo

            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveCapNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalFreeSpaceNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveUsageNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveCapNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveFreeSpaceNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighCPUTimeNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighMemoryTimeNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighIOTimeNow, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveCap7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalFreeSpace7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveUsage7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveCap7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveFreeSpace7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighCPUTime7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighMemoryTime7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighIOTime7Days, null), "Text", "");

            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalFreeSpaceNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpaceNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalDriveUsageNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalDriveUsageNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveCapNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveCapNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveFreeSpaceNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveFreeSpaceNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighCPUTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighCPUTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighMemoryTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighMemoryTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighIOTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighIOTimeNow, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpace7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpace7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalDriveUsage7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalDriveUsage7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveCap7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveCap7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveFreeSpace7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveFreeSpace7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighCPUTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighCPUTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighMemoryTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighMemoryTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighIOTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighIOTime7Days, null), "ShowControl", (Int16)3);

            // GroupBox Software OS Health

            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSVersionArchitecture, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWMIStatus, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthLastSystemUpdate, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWindowsUpdateStatus, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfApps, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfExes, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfBins, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSEndOfSupport, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSIE11Support, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWin10Ready, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSComplience, null), "Text", "");

            // GroupBox Security Compliance

            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceInetSecuritySettings, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceUserAccountStatus, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusRTP, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusUpToDate, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusNumber, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusAll, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareRTP, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareUpToDate, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareNumber, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareAll, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallName, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallRTP, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallNumber, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallAll, null), "Text", "");

            // GroupBox Network Health
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingNetTaffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingNetTaffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalNetTaffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessNetConnectionRatio24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthNetAvailLevel24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingNetBitrate24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgNetResponseTime24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingWebTraffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingWebTraffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalWebTraffic24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingWebBitrate24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebRequestSize24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebResponseSize24Hours, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours, null), "Text", "");

            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingNetTaffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingNetTaffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalNetTaffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessNetConnectionRatio7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthNetAvailLevel7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingNetBitrate7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingNetBitrate7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgNetResponseTime7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingWebTraffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingWebTraffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalWebTraffic7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingWebBitrate7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingWebBitrate7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebRequestSize7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebResponseSize7Days, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessHTTPRequestRatio7Days, null), "Text", "");

            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalNetTaffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalNetTaffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessNetConnectionRatio24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessNetConnectionRatio7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessNetConnectionRatio7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthNetAvailLevel24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthNetAvailLevel24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthNetAvailLevel7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthNetAvailLevel7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingNetBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingNetBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingNetBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingNetBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgNetResponseTime24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgNetResponseTime24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgNetResponseTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgNetResponseTime7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalWebTraffic7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingWebBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingWebBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingWebBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingWebBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebRequestSize24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebRequestSize24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebRequestSize7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebRequestSize7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebResponseSize24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebResponseSize24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebResponseSize7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebResponseSize7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessHTTPRequestRatio7Days, null), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days, null), "ShowControl", (Int16)3);

            // GroupBox L1-Checkliste

            _.SETm1a0(this, _.NnO(_env.TextBoxL1FreeSpace, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1OSUpToDate, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Browser, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Collaboration, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus2, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus3, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Defender, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1BootLogon2, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1BootLogon3, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1CPU24, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1CPU7, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Speicher24, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Speicher7, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Bluescreen24, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Bluescrren7, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1HardReset24, null), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1HardReset7, null), "Text", "");

            // --- GroupBoxGeneralInfo

            nexthinkBaseURL = _.CONCAT(_.CALLm1argp(this, _outer, "getNexthinkBaseURL", _.ARGS.ForceBrackets()), "query?p1=");
            nexthinkQuery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml";

            colorWarning = "#F20012";
            colorCheck = "#1B709F";

            hostname = _.VAL(_.CALLm1v5(this, _.NnO(_env.hlobj, null), "GetValue", "ComputerDetail.Hostname", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            if (_.IF(_.EQ(_.NullableSTR(hostname), "")))
            {
                _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", "Der Computer hat keinen Hostnamen.");
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            nexthinkURL = _.CONCAT(nexthinkBaseURL, _.UCASE(hostname), nexthinkQuery);
            nexthinkURL = "https://httpbin.org/get";

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);

            //MsgBox nexthinkURL

            //time of call
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCallTime, null), "Text", _.FORMATDATETIME(_.NOW(), VBScriptConstants.vbGeneralDate));
            });

            _.HANDLEERROR(errOn, () => {
                xmlhttp = _.CREATEOBJECT("Msxml2.ServerXMLHTTP.6.0");
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v2(this, _.NnO(xmlhttp, "xmlhttp"), "setOption", (Int16)2, (Int16)13056);
            }); //bypass certificate errors
            _.HANDLEERROR(errOn, () => {
                _.CALLm1argp(this, _.NnO(xmlhttp, "xmlhttp"), "open", _.ARGS.Val("GET").Ref(nexthinkURL, v => { nexthinkURL = v; }).Val(false).Val(_.CALLm1argp(this, _outer, "getNexthinkUser", _.ARGS.ForceBrackets())).Val(_.CALLm1argp(this, _outer, "getNexthinkPassword", _.ARGS.ForceBrackets())));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v0(this, _.NnO(xmlhttp, "xmlhttp"), "send");
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, null), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", "Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "HelpContext")));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            //Reset the Error Data
            _.HANDLEERROR(errOn, () => {
                _.CLEARANYERROR();
            });

            _.HANDLEERROR(errOn, () => {
                xmlDoc = _.CREATEOBJECT("Msxml2.DOMDocument");
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(xmlDoc, "xmlDoc"), "async", "false");
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(xmlDoc, "xmlDoc"), "load", _.CALLm1v0(this, _.NnO(xmlhttp, "xmlhttp"), "responseXML"));
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, null), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", "Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.NnO(_.ERR, null), "HelpContext")));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            _.HANDLEERROR(errOn, () => {
                dict = _.CREATEOBJECT("Scripting.Dictionary");
            });

            //iterate all nodes and write into dictionary
            _.HANDLEERROR(errOn, () => {
                i = (Int16)0;
            });
            IEnumerator enumerationContent = null;
            _.HANDLEERROR(errOn, () => {
                enumerationContent = _.ENUMERABLE(_.CALLm1v1(this, _.NnO(xmlDoc, "xmlDoc"), "SelectNodes", "//table/header/*")).GetEnumerator();
            });
            while (true)
            {
                if (enumerationContent != null)
                {
                    if (!enumerationContent.MoveNext())
                    {
                        n = null;
                        break;
                    }
                    n = enumerationContent.Current;
                }
                _.HANDLEERROR(errOn, () => {
                    curnode = _.OBJ(_.CALLm2v1(this, _.NnO(xmlDoc, "xmlDoc"), "documentElement", "selectSingleNode", _.CONCAT("//table/body/r/c", i)));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v2(this, _.NnO(dict, "dict"), "Add", _.CALLm1v0(this, _.NnO(n, "n"), "Text"), _.CALLm1v0(this, _.NnO(curnode, "curnode"), "Text"));
                });
                _.HANDLEERROR(errOn, () => {
                    i = _.ADD(i, (Int16)1);
                });
                if (enumerationContent == null)
                    break;
            }

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, null), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, null), "MsgBox", "Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten.");
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            // from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

            //fill textboxes
            //LabelNName.Text = dict.key("name")
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralHostName, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastIP, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_ip_address")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxlGeneralDeviceManufacturer, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "device_manufacturer")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceProductVersion, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "device_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralOS, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "os_version_and_architecture")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGroupName, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "group_name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastBootTime, null), "Text", _.FORMATDATETIME(_.REPLACE(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_system_boot"), "T", " "), VBScriptConstants.vbGeneralDate));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastLogon, null), "Text", _.FORMATDATETIME(_.REPLACE(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_logon_time"), "T", " "), VBScriptConstants.vbGeneralDate));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceType, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_logged_on_user")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralBIOSSerialNumber, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "bios_serial_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUModel, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "cpu_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCPUs, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_cpus")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfLogProcs, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "logical_cpu_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCores, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_cores")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUFreq, null), "Text", _.CONCAT(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "cpu_frequency"), " MHz"));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralTotalRAM, null), "Text", _.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "total_ram"))));
            });

            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfGraphCards, null), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_graphical_cards")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGraphCardRAM, null), "Text", _.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "graphical_card_ram"))));
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
