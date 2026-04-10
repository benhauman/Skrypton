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

            _.SETm1a0(this, _.NnO(_env.TabPageGeneralInfo, "TabPageGeneralInfo"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.TabPageSoftwareOSHealth, "TabPageSoftwareOSHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageSecurityCompliance, "TabPageSecurityCompliance"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageTechnicalInfo, "TabPageTechnicalInfo"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageNetworkHealth, "TabPageNetworkHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.TabPageL1Checklist, "TabPageL1Checklist"), "ShowControl", (Int16)3);

            _.SETm1a0(this, _.NnO(_env.GroupBoxGeneralInfo, "GroupBoxGeneralInfo"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.GroupBoxTechnicalInfo, "GroupBoxTechnicalInfo"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxSoftwareOSHealth, "GroupBoxSoftwareOSHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxSecurityCompliance, "GroupBoxSecurityCompliance"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxNetworkHealth, "GroupBoxNetworkHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.GroupBoxL1Checklist, "GroupBoxL1Checklist"), "ShowControl", (Int16)3);

            _.SETm1a0(this, _.NnO(_env.ButtonGeneralInfo, "ButtonGeneralInfo"), "BackColor", "#5b5b5b");
            _.SETm1a0(this, _.NnO(_env.ButtonTechnicalInfo, "ButtonTechnicalInfo"), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonSWHealth, "ButtonSWHealth"), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonSecurityCompliance, "ButtonSecurityCompliance"), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonNetworkHealth, "ButtonNetworkHealth"), "BackColor", "#1B709F");
            _.SETm1a0(this, _.NnO(_env.ButtonL1Checklist, "ButtonL1Checklist"), "BackColor", "#1B709F");

            _.SETm1a0(this, _.NnO(_env.TabControlNexthink, "TabControlNexthink"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _.NnO(_env.TabPageGeneralInfo, "TabPageGeneralInfo"), "RequestFocus", true);

            //Clear TextBoxes

            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCallTime, "TextBoxGeneralCallTime"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralHostName, "TextBoxGeneralHostName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxlGeneralDeviceManufacturer, "TextBoxlGeneralDeviceManufacturer"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceProductVersion, "TextBoxGeneralDeviceProductVersion"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastIP, "TextBoxGeneralLastIP"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGroupName, "TextBoxGeneralGroupName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralOS, "TextBoxGeneralOS"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastBootTime, "TextBoxGeneralLastBootTime"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastLogon, "TextBoxGeneralLastLogon"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceType, "TextBoxGeneralDeviceType"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralBIOSSerialNumber, "TextBoxGeneralBIOSSerialNumber"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUModel, "TextBoxGeneralCPUModel"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCPUs, "TextBoxGeneralNumberOfCPUs"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfLogProcs, "TextBoxGeneralNumberOfLogProcs"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCores, "TextBoxGeneralNumberOfCores"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUFreq, "TextBoxGeneralCPUFreq"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralTotalRAM, "TextBoxGeneralTotalRAM"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfGraphCards, "TextBoxGeneralNumberOfGraphCards"), "Text", "");

            // --- GroupBoxTechnicalInfo

            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveCapNow, "TextBoxTechnicalInfoTotalDriveCapNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalFreeSpaceNow, "TextBoxTechnicalInfoTotalFreeSpaceNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveUsageNow, "TextBoxTechnicalInfoTotalDriveUsageNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveCapNow, "TextBoxTechnicalInfoSystemDriveCapNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveFreeSpaceNow, "TextBoxTechnicalInfoSystemDriveFreeSpaceNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighCPUTimeNow, "TextBoxTechnicalInfoHighCPUTimeNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighMemoryTimeNow, "TextBoxTechnicalInfoHighMemoryTimeNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighIOTimeNow, "TextBoxTechnicalInfoHighIOTimeNow"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveCap7Days, "TextBoxTechnicalInfoTotalDriveCap7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalFreeSpace7Days, "TextBoxTechnicalInfoTotalFreeSpace7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoTotalDriveUsage7Days, "TextBoxTechnicalInfoTotalDriveUsage7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveCap7Days, "TextBoxTechnicalInfoSystemDriveCap7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoSystemDriveFreeSpace7Days, "TextBoxTechnicalInfoSystemDriveFreeSpace7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighCPUTime7Days, "TextBoxTechnicalInfoHighCPUTime7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighMemoryTime7Days, "TextBoxTechnicalInfoHighMemoryTime7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxTechnicalInfoHighIOTime7Days, "TextBoxTechnicalInfoHighIOTime7Days"), "Text", "");

            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalFreeSpaceNow, "ImageNOKTechnicalInfoTotalFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpaceNow, "ImageOKTechnicalInfoTotalFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalDriveUsageNow, "ImageNOKTechnicalInfoTotalDriveUsageNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalDriveUsageNow, "ImageOKTechnicalInfoTotalDriveUsageNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveCapNow, "ImageNOKTechnicalInfoSystemDriveCapNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveCapNow, "ImageOKTechnicalInfoSystemDriveCapNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveFreeSpaceNow, "ImageNOKTechnicalInfoSystemDriveFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveFreeSpaceNow, "ImageOKTechnicalInfoSystemDriveFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighCPUTimeNow, "ImageNOKTechnicalInfoHighCPUTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighCPUTimeNow, "ImageOKTechnicalInfoHighCPUTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighMemoryTimeNow, "ImageNOKTechnicalInfoHighMemoryTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighMemoryTimeNow, "ImageOKTechnicalInfoHighMemoryTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighIOTimeNow, "ImageNOKTechnicalInfoHighIOTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighIOTimeNow, "ImageOKTechnicalInfoHighIOTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpace7Days, "ImageOKTechnicalInfoTotalFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalFreeSpace7Days, "ImageOKTechnicalInfoTotalFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoTotalDriveUsage7Days, "ImageNOKTechnicalInfoTotalDriveUsage7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoTotalDriveUsage7Days, "ImageOKTechnicalInfoTotalDriveUsage7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveCap7Days, "ImageNOKTechnicalInfoSystemDriveCap7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveCap7Days, "ImageOKTechnicalInfoSystemDriveCap7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoSystemDriveFreeSpace7Days, "ImageNOKTechnicalInfoSystemDriveFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoSystemDriveFreeSpace7Days, "ImageOKTechnicalInfoSystemDriveFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighCPUTime7Days, "ImageNOKTechnicalInfoHighCPUTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighCPUTime7Days, "ImageOKTechnicalInfoHighCPUTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighMemoryTime7Days, "ImageNOKTechnicalInfoHighMemoryTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighMemoryTime7Days, "ImageOKTechnicalInfoHighMemoryTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKTechnicalInfoHighIOTime7Days, "ImageNOKTechnicalInfoHighIOTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKTechnicalInfoHighIOTime7Days, "ImageOKTechnicalInfoHighIOTime7Days"), "ShowControl", (Int16)3);

            // GroupBox Software OS Health

            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSVersionArchitecture, "TextBoxSoftwareOSHealthOSVersionArchitecture"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSName, "TextBoxSoftwareOSHealthOSName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWMIStatus, "TextBoxSoftwareOSHealthWMIStatus"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthLastSystemUpdate, "TextBoxSoftwareOSHealthLastSystemUpdate"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWindowsUpdateStatus, "TextBoxSoftwareOSHealthWindowsUpdateStatus"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfApps, "TextBoxSoftwareOSHealthNumberOfApps"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfExes, "TextBoxSoftwareOSHealthNumberOfExes"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthNumberOfBins, "TextBoxSoftwareOSHealthNumberOfBins"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSEndOfSupport, "TextBoxSoftwareOSHealthOSEndOfSupport"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSIE11Support, "TextBoxSoftwareOSHealthOSIE11Support"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthWin10Ready, "TextBoxSoftwareOSHealthWin10Ready"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSoftwareOSHealthOSComplience, "TextBoxSoftwareOSHealthOSComplience"), "Text", "");

            // GroupBox Security Compliance

            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceInetSecuritySettings, "TextBoxSecurityComplianceInetSecuritySettings"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceUserAccountStatus, "TextBoxSecurityComplianceUserAccountStatus"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusName, "TextBoxSecurityComplianceAntivirusName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusRTP, "TextBoxSecurityComplianceAntivirusRTP"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusUpToDate, "TextBoxSecurityComplianceAntivirusUpToDate"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusNumber, "TextBoxSecurityComplianceAntivirusNumber"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntivirusAll, "TextBoxSecurityComplianceAntivirusAll"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareName, "TextBoxSecurityComplianceAntispywareName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareRTP, "TextBoxSecurityComplianceAntispywareRTP"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareUpToDate, "TextBoxSecurityComplianceAntispywareUpToDate"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareNumber, "TextBoxSecurityComplianceAntispywareNumber"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceAntispywareAll, "TextBoxSecurityComplianceAntispywareAll"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallName, "TextBoxSecurityComplianceFirewallName"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallRTP, "TextBoxSecurityComplianceFirewallRTP"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallNumber, "TextBoxSecurityComplianceFirewallNumber"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxSecurityComplianceFirewallAll, "TextBoxSecurityComplianceFirewallAll"), "Text", "");

            // GroupBox Network Health
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingNetTaffic24Hours, "TextBoxNetworkHealthIncomingNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingNetTaffic24Hours, "TextBoxNetworkHealthOutgoingNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalNetTaffic24Hours, "TextBoxNetworkHealthTotalNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessNetConnectionRatio24Hours, "TextBoxNetworkHealthSuccessNetConnectionRatio24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthNetAvailLevel24Hours, "TextBoxNetworkHealthNetAvailLevel24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingNetBitrate24Hours, "TextBoxNetworkHealthAvgIncomingNetBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours, "TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgNetResponseTime24Hours, "TextBoxNetworkHealthAvgNetResponseTime24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingWebTraffic24Hours, "TextBoxNetworkHealthIncomingWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingWebTraffic24Hours, "TextBoxNetworkHealthOutgoingWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalWebTraffic24Hours, "TextBoxNetworkHealthTotalWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingWebBitrate24Hours, "TextBoxNetworkHealthAvgIncomingWebBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours, "TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebRequestSize24Hours, "TextBoxNetworkHealthAvgWebRequestSize24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebResponseSize24Hours, "TextBoxNetworkHealthAvgWebResponseSize24Hours"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours, "TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours"), "Text", "");

            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingNetTaffic7Days, "TextBoxNetworkHealthIncomingNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingNetTaffic7Days, "TextBoxNetworkHealthOutgoingNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalNetTaffic7Days, "TextBoxNetworkHealthTotalNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessNetConnectionRatio7Days, "TextBoxNetworkHealthSuccessNetConnectionRatio7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthNetAvailLevel7Days, "TextBoxNetworkHealthNetAvailLevel7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingNetBitrate7Days, "TextBoxNetworkHealthAvgIncomingNetBitrate7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingNetBitrate7Days, "TextBoxNetworkHealthAvgOutgoingNetBitrate7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgNetResponseTime7Days, "TextBoxNetworkHealthAvgNetResponseTime7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthIncomingWebTraffic7Days, "TextBoxNetworkHealthIncomingWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthOutgoingWebTraffic7Days, "TextBoxNetworkHealthOutgoingWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthTotalWebTraffic7Days, "TextBoxNetworkHealthTotalWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgIncomingWebBitrate7Days, "TextBoxNetworkHealthAvgIncomingWebBitrate7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgOutgoingWebBitrate7Days, "TextBoxNetworkHealthAvgOutgoingWebBitrate7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebRequestSize7Days, "TextBoxNetworkHealthAvgWebRequestSize7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthAvgWebResponseSize7Days, "TextBoxNetworkHealthAvgWebResponseSize7Days"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxNetworkHealthSuccessHTTPRequestRatio7Days, "TextBoxNetworkHealthSuccessHTTPRequestRatio7Days"), "Text", "");

            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingNetTaffic24Hours, "ImageOKNetworkHealthIncomingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingNetTaffic24Hours, "ImageNOKNetworkHealthIncomingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingNetTaffic7Days, "ImageOKNetworkHealthIncomingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingNetTaffic7Days, "ImageNOKNetworkHealthIncomingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingNetTaffic24Hours, "ImageOKNetworkHealthOutgoingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingNetTaffic24Hours, "ImageNOKNetworkHealthOutgoingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingNetTaffic7Days, "ImageOKNetworkHealthOutgoingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingNetTaffic7Days, "ImageNOKNetworkHealthOutgoingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalNetTaffic24Hours, "ImageOKNetworkHealthTotalNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalNetTaffic24Hours, "ImageNOKNetworkHealthTotalNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalNetTaffic7Days, "ImageOKNetworkHealthTotalNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalNetTaffic7Days, "ImageNOKNetworkHealthTotalNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessNetConnectionRatio24Hours, "ImageOKNetworkHealthSuccessNetConnectionRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours, "ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessNetConnectionRatio7Days, "ImageOKNetworkHealthSuccessNetConnectionRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessNetConnectionRatio7Days, "ImageNOKNetworkHealthSuccessNetConnectionRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthNetAvailLevel24Hours, "ImageOKNetworkHealthNetAvailLevel24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthNetAvailLevel24Hours, "ImageNOKNetworkHealthNetAvailLevel24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthNetAvailLevel7Days, "ImageOKNetworkHealthNetAvailLevel7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthNetAvailLevel7Days, "ImageNOKNetworkHealthNetAvailLevel7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingNetBitrate24Hours, "ImageOKNetworkHealthAvgIncomingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours, "ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingNetBitrate7Days, "ImageOKNetworkHealthAvgIncomingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingNetBitrate7Days, "ImageNOKNetworkHealthAvgIncomingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours, "ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours, "ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingNetBitrate7Days, "ImageOKNetworkHealthAvgOutgoingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days, "ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgNetResponseTime24Hours, "ImageOKNetworkHealthAvgNetResponseTime24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgNetResponseTime24Hours, "ImageNOKNetworkHealthAvgNetResponseTime24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgNetResponseTime7Days, "ImageOKNetworkHealthAvgNetResponseTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgNetResponseTime7Days, "ImageNOKNetworkHealthAvgNetResponseTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingWebTraffic24Hours, "ImageOKNetworkHealthIncomingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingWebTraffic24Hours, "ImageNOKNetworkHealthIncomingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthIncomingWebTraffic7Days, "ImageOKNetworkHealthIncomingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthIncomingWebTraffic7Days, "ImageNOKNetworkHealthIncomingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingWebTraffic24Hours, "ImageOKNetworkHealthOutgoingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingWebTraffic24Hours, "ImageNOKNetworkHealthOutgoingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthOutgoingWebTraffic7Days, "ImageOKNetworkHealthOutgoingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthOutgoingWebTraffic7Days, "ImageNOKNetworkHealthOutgoingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic24Hours, "ImageOKNetworkHealthTotalWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic24Hours, "ImageOKNetworkHealthTotalWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthTotalWebTraffic7Days, "ImageOKNetworkHealthTotalWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthTotalWebTraffic7Days, "ImageNOKNetworkHealthTotalWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingWebBitrate24Hours, "ImageOKNetworkHealthAvgIncomingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours, "ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgIncomingWebBitrate7Days, "ImageOKNetworkHealthAvgIncomingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgIncomingWebBitrate7Days, "ImageNOKNetworkHealthAvgIncomingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours, "ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours, "ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgOutgoingWebBitrate7Days, "ImageOKNetworkHealthAvgOutgoingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days, "ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebRequestSize24Hours, "ImageOKNetworkHealthAvgWebRequestSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebRequestSize24Hours, "ImageNOKNetworkHealthAvgWebRequestSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebRequestSize7Days, "ImageOKNetworkHealthAvgWebRequestSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebRequestSize7Days, "ImageNOKNetworkHealthAvgWebRequestSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebResponseSize24Hours, "ImageOKNetworkHealthAvgWebResponseSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebResponseSize24Hours, "ImageNOKNetworkHealthAvgWebResponseSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthAvgWebResponseSize7Days, "ImageOKNetworkHealthAvgWebResponseSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthAvgWebResponseSize7Days, "ImageNOKNetworkHealthAvgWebResponseSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours, "ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours, "ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageOKNetworkHealthSuccessHTTPRequestRatio7Days, "ImageOKNetworkHealthSuccessHTTPRequestRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _.NnO(_env.ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days, "ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days"), "ShowControl", (Int16)3);

            // GroupBox L1-Checkliste

            _.SETm1a0(this, _.NnO(_env.TextBoxL1FreeSpace, "TextBoxL1FreeSpace"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1OSUpToDate, "TextBoxL1OSUpToDate"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Browser, "TextBoxL1Browser"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Collaboration, "TextBoxL1Collaboration"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus, "TextBoxL1Antivirus"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus2, "TextBoxL1Antivirus2"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Antivirus3, "TextBoxL1Antivirus3"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Defender, "TextBoxL1Defender"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1BootLogon2, "TextBoxL1BootLogon2"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1BootLogon3, "TextBoxL1BootLogon3"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1CPU24, "TextBoxL1CPU24"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1CPU7, "TextBoxL1CPU7"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Speicher24, "TextBoxL1Speicher24"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Speicher7, "TextBoxL1Speicher7"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Bluescreen24, "TextBoxL1Bluescreen24"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1Bluescrren7, "TextBoxL1Bluescrren7"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1HardReset24, "TextBoxL1HardReset24"), "Text", "");
            _.SETm1a0(this, _.NnO(_env.TextBoxL1HardReset7, "TextBoxL1HardReset7"), "Text", "");

            // --- GroupBoxGeneralInfo

            nexthinkBaseURL = _.CONCAT(_.CALLm1argp(this, _outer, "getNexthinkBaseURL", _.ARGS.ForceBrackets()), "query?p1=");
            nexthinkQuery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml";

            colorWarning = "#F20012";
            colorCheck = "#1B709F";

            hostname = _.VAL(_.CALLm1v5(this, _.NnO(_env.hlobj, "hlobj"), "GetValue", "ComputerDetail.Hostname", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            if (_.IF(_.EQ(_.NullableSTR(hostname), "")))
            {
                _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", "Der Computer hat keinen Hostnamen.");
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            nexthinkURL = _.CONCAT(nexthinkBaseURL, _.UCASE(hostname), nexthinkQuery);
            nexthinkURL = "https://httpbin.org/get";

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);

            //MsgBox nexthinkURL

            //time of call
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCallTime, "TextBoxGeneralCallTime"), "Text", _.FORMATDATETIME(_.NOW(), VBScriptConstants.vbGeneralDate));
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
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", "Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "HelpContext")));
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
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", "Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "HelpContext")));
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
                        break;
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
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.NnO(_.ERR, "ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _.NnO(_env.model, "model"), "MsgBox", "Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten.");
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            // from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

            //fill textboxes
            //LabelNName.Text = dict.key("name")
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralHostName, "TextBoxGeneralHostName"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastIP, "TextBoxGeneralLastIP"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_ip_address")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxlGeneralDeviceManufacturer, "TextBoxlGeneralDeviceManufacturer"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "device_manufacturer")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceProductVersion, "TextBoxGeneralDeviceProductVersion"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "device_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralOS, "TextBoxGeneralOS"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "os_version_and_architecture")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGroupName, "TextBoxGeneralGroupName"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "group_name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastBootTime, "TextBoxGeneralLastBootTime"), "Text", _.FORMATDATETIME(_.REPLACE(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_system_boot"), "T", " "), VBScriptConstants.vbGeneralDate));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralLastLogon, "TextBoxGeneralLastLogon"), "Text", _.FORMATDATETIME(_.REPLACE(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_logon_time"), "T", " "), VBScriptConstants.vbGeneralDate));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralDeviceType, "TextBoxGeneralDeviceType"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "last_logged_on_user")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralBIOSSerialNumber, "TextBoxGeneralBIOSSerialNumber"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "bios_serial_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUModel, "TextBoxGeneralCPUModel"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "cpu_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCPUs, "TextBoxGeneralNumberOfCPUs"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_cpus")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfLogProcs, "TextBoxGeneralNumberOfLogProcs"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "logical_cpu_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfCores, "TextBoxGeneralNumberOfCores"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_cores")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralCPUFreq, "TextBoxGeneralCPUFreq"), "Text", _.CONCAT(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "cpu_frequency"), " MHz"));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralTotalRAM, "TextBoxGeneralTotalRAM"), "Text", _.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "total_ram"))));
            });

            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralNumberOfGraphCards, "TextBoxGeneralNumberOfGraphCards"), "Text", _.VAL(_.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "number_of_graphical_cards")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _.NnO(_env.TextBoxGeneralGraphCardRAM, "TextBoxGeneralGraphCardRAM"), "Text", _.VAL(_.CALLm1v1(this, _outer, "ConvertSize", _.CALLm1v1(this, _.NnO(dict, "dict"), "Item", "graphical_card_ram"))));
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
