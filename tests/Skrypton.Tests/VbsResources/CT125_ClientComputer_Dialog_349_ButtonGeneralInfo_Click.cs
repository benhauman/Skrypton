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
            _.CALLm1v0(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ButtonGeneralInfo_Click");
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
                Size = _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ROUND", _.DIV(Size, (Int16)1024), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " MB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ROUND", _.DIV(Size, 1048576), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " GB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ROUND", _.DIV(Size, 1073741824), (Int16)2));
            }
            else if (_.IF(_.EQ(Suffix, " TB")))
            {
                Size = _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "ROUND", _.DIV(Size, 1099511627776d), (Int16)2));
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

            _.SETm1a0(this, _env.TabPageGeneralInfo ?? throw new InvalidOperationException("Reference not set:TabPageGeneralInfo"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _env.TabPageSoftwareOSHealth ?? throw new InvalidOperationException("Reference not set:TabPageSoftwareOSHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.TabPageSecurityCompliance ?? throw new InvalidOperationException("Reference not set:TabPageSecurityCompliance"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.TabPageTechnicalInfo ?? throw new InvalidOperationException("Reference not set:TabPageTechnicalInfo"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.TabPageNetworkHealth ?? throw new InvalidOperationException("Reference not set:TabPageNetworkHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.TabPageL1Checklist ?? throw new InvalidOperationException("Reference not set:TabPageL1Checklist"), "ShowControl", (Int16)3);

            _.SETm1a0(this, _env.GroupBoxGeneralInfo ?? throw new InvalidOperationException("Reference not set:GroupBoxGeneralInfo"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _env.GroupBoxTechnicalInfo ?? throw new InvalidOperationException("Reference not set:GroupBoxTechnicalInfo"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.GroupBoxSoftwareOSHealth ?? throw new InvalidOperationException("Reference not set:GroupBoxSoftwareOSHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.GroupBoxSecurityCompliance ?? throw new InvalidOperationException("Reference not set:GroupBoxSecurityCompliance"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.GroupBoxNetworkHealth ?? throw new InvalidOperationException("Reference not set:GroupBoxNetworkHealth"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.GroupBoxL1Checklist ?? throw new InvalidOperationException("Reference not set:GroupBoxL1Checklist"), "ShowControl", (Int16)3);

            _.SETm1a0(this, _env.ButtonGeneralInfo ?? throw new InvalidOperationException("Reference not set:ButtonGeneralInfo"), "BackColor", "#5b5b5b");
            _.SETm1a0(this, _env.ButtonTechnicalInfo ?? throw new InvalidOperationException("Reference not set:ButtonTechnicalInfo"), "BackColor", "#1B709F");
            _.SETm1a0(this, _env.ButtonSWHealth ?? throw new InvalidOperationException("Reference not set:ButtonSWHealth"), "BackColor", "#1B709F");
            _.SETm1a0(this, _env.ButtonSecurityCompliance ?? throw new InvalidOperationException("Reference not set:ButtonSecurityCompliance"), "BackColor", "#1B709F");
            _.SETm1a0(this, _env.ButtonNetworkHealth ?? throw new InvalidOperationException("Reference not set:ButtonNetworkHealth"), "BackColor", "#1B709F");
            _.SETm1a0(this, _env.ButtonL1Checklist ?? throw new InvalidOperationException("Reference not set:ButtonL1Checklist"), "BackColor", "#1B709F");

            _.SETm1a0(this, _env.TabControlNexthink ?? throw new InvalidOperationException("Reference not set:TabControlNexthink"), "ShowControl", (Int16)1);
            _.SETm1a0(this, _env.TabPageGeneralInfo ?? throw new InvalidOperationException("Reference not set:TabPageGeneralInfo"), "RequestFocus", true);

            //Clear TextBoxes

            _.SETm1a0(this, _env.TextBoxGeneralCallTime ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCallTime"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralHostName ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralHostName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxlGeneralDeviceManufacturer ?? throw new InvalidOperationException("Reference not set:TextBoxlGeneralDeviceManufacturer"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralDeviceProductVersion ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralDeviceProductVersion"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralLastIP ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastIP"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralGroupName ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralGroupName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralOS ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralOS"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralLastBootTime ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastBootTime"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralLastLogon ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastLogon"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralDeviceType ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralDeviceType"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralBIOSSerialNumber ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralBIOSSerialNumber"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralCPUModel ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCPUModel"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralNumberOfCPUs ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfCPUs"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralNumberOfLogProcs ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfLogProcs"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralNumberOfCores ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfCores"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralCPUFreq ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCPUFreq"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralTotalRAM ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralTotalRAM"), "Text", "");
            _.SETm1a0(this, _env.TextBoxGeneralNumberOfGraphCards ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfGraphCards"), "Text", "");

            // --- GroupBoxTechnicalInfo

            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalDriveCapNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalDriveCapNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalFreeSpaceNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalDriveUsageNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalDriveUsageNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoSystemDriveCapNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoSystemDriveCapNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoSystemDriveFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoSystemDriveFreeSpaceNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighCPUTimeNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighCPUTimeNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighMemoryTimeNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighMemoryTimeNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighIOTimeNow ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighIOTimeNow"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalDriveCap7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalDriveCap7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalFreeSpace7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoTotalDriveUsage7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoTotalDriveUsage7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoSystemDriveCap7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoSystemDriveCap7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoSystemDriveFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoSystemDriveFreeSpace7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighCPUTime7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighCPUTime7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighMemoryTime7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighMemoryTime7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxTechnicalInfoHighIOTime7Days ?? throw new InvalidOperationException("Reference not set:TextBoxTechnicalInfoHighIOTime7Days"), "Text", "");

            _.SETm1a0(this, _env.ImageNOKTechnicalInfoTotalFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoTotalFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoTotalFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoTotalFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoTotalDriveUsageNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoTotalDriveUsageNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoTotalDriveUsageNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoTotalDriveUsageNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoSystemDriveCapNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoSystemDriveCapNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoSystemDriveCapNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoSystemDriveCapNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoSystemDriveFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoSystemDriveFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoSystemDriveFreeSpaceNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoSystemDriveFreeSpaceNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighCPUTimeNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighCPUTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighCPUTimeNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighCPUTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighMemoryTimeNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighMemoryTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighMemoryTimeNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighMemoryTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighIOTimeNow ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighIOTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighIOTimeNow ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighIOTimeNow"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoTotalFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoTotalFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoTotalFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoTotalFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoTotalDriveUsage7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoTotalDriveUsage7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoTotalDriveUsage7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoTotalDriveUsage7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoSystemDriveCap7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoSystemDriveCap7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoSystemDriveCap7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoSystemDriveCap7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoSystemDriveFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoSystemDriveFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoSystemDriveFreeSpace7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoSystemDriveFreeSpace7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighCPUTime7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighCPUTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighCPUTime7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighCPUTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighMemoryTime7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighMemoryTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighMemoryTime7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighMemoryTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKTechnicalInfoHighIOTime7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKTechnicalInfoHighIOTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKTechnicalInfoHighIOTime7Days ?? throw new InvalidOperationException("Reference not set:ImageOKTechnicalInfoHighIOTime7Days"), "ShowControl", (Int16)3);

            // GroupBox Software OS Health

            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthOSVersionArchitecture ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthOSVersionArchitecture"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthOSName ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthOSName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthWMIStatus ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthWMIStatus"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthLastSystemUpdate ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthLastSystemUpdate"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthWindowsUpdateStatus ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthWindowsUpdateStatus"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthNumberOfApps ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthNumberOfApps"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthNumberOfExes ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthNumberOfExes"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthNumberOfBins ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthNumberOfBins"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthOSEndOfSupport ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthOSEndOfSupport"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthOSIE11Support ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthOSIE11Support"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthWin10Ready ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthWin10Ready"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSoftwareOSHealthOSComplience ?? throw new InvalidOperationException("Reference not set:TextBoxSoftwareOSHealthOSComplience"), "Text", "");

            // GroupBox Security Compliance

            _.SETm1a0(this, _env.TextBoxSecurityComplianceInetSecuritySettings ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceInetSecuritySettings"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceUserAccountStatus ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceUserAccountStatus"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntivirusName ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntivirusName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntivirusRTP ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntivirusRTP"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntivirusUpToDate ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntivirusUpToDate"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntivirusNumber ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntivirusNumber"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntivirusAll ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntivirusAll"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntispywareName ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntispywareName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntispywareRTP ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntispywareRTP"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntispywareUpToDate ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntispywareUpToDate"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntispywareNumber ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntispywareNumber"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceAntispywareAll ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceAntispywareAll"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceFirewallName ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceFirewallName"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceFirewallRTP ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceFirewallRTP"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceFirewallNumber ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceFirewallNumber"), "Text", "");
            _.SETm1a0(this, _env.TextBoxSecurityComplianceFirewallAll ?? throw new InvalidOperationException("Reference not set:TextBoxSecurityComplianceFirewallAll"), "Text", "");

            // GroupBox Network Health
            _.SETm1a0(this, _env.TextBoxNetworkHealthIncomingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthIncomingNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthOutgoingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthOutgoingNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthTotalNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthTotalNetTaffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthSuccessNetConnectionRatio24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthSuccessNetConnectionRatio24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthNetAvailLevel24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthNetAvailLevel24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgIncomingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgIncomingNetBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgNetResponseTime24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgNetResponseTime24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthIncomingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthIncomingWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthOutgoingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthOutgoingWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthTotalWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthTotalWebTraffic24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgIncomingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgIncomingWebBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgWebRequestSize24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgWebRequestSize24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgWebResponseSize24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgWebResponseSize24Hours"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours"), "Text", "");

            _.SETm1a0(this, _env.TextBoxNetworkHealthIncomingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthIncomingNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthOutgoingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthOutgoingNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthTotalNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthTotalNetTaffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthSuccessNetConnectionRatio7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthSuccessNetConnectionRatio7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthNetAvailLevel7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthNetAvailLevel7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgIncomingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgIncomingNetBitrate7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgOutgoingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgOutgoingNetBitrate7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgNetResponseTime7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgNetResponseTime7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthIncomingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthIncomingWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthOutgoingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthOutgoingWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthTotalWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthTotalWebTraffic7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgIncomingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgIncomingWebBitrate7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgOutgoingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgOutgoingWebBitrate7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgWebRequestSize7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgWebRequestSize7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthAvgWebResponseSize7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthAvgWebResponseSize7Days"), "Text", "");
            _.SETm1a0(this, _env.TextBoxNetworkHealthSuccessHTTPRequestRatio7Days ?? throw new InvalidOperationException("Reference not set:TextBoxNetworkHealthSuccessHTTPRequestRatio7Days"), "Text", "");

            _.SETm1a0(this, _env.ImageOKNetworkHealthIncomingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthIncomingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthIncomingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthIncomingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthIncomingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthIncomingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthIncomingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthIncomingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthOutgoingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthOutgoingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthOutgoingNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthOutgoingNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthOutgoingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthOutgoingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthOutgoingNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthOutgoingNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthTotalNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthTotalNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthTotalNetTaffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthTotalNetTaffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthTotalNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthTotalNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthTotalNetTaffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthTotalNetTaffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthSuccessNetConnectionRatio24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthSuccessNetConnectionRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthSuccessNetConnectionRatio7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthSuccessNetConnectionRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthSuccessNetConnectionRatio7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthSuccessNetConnectionRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthNetAvailLevel24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthNetAvailLevel24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthNetAvailLevel24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthNetAvailLevel24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthNetAvailLevel7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthNetAvailLevel7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthNetAvailLevel7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthNetAvailLevel7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgIncomingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgIncomingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgIncomingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgIncomingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgIncomingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgIncomingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgOutgoingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgOutgoingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgNetResponseTime24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgNetResponseTime24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgNetResponseTime24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgNetResponseTime24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgNetResponseTime7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgNetResponseTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgNetResponseTime7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgNetResponseTime7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthIncomingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthIncomingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthIncomingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthIncomingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthIncomingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthIncomingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthIncomingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthIncomingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthOutgoingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthOutgoingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthOutgoingWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthOutgoingWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthOutgoingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthOutgoingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthOutgoingWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthOutgoingWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthTotalWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthTotalWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthTotalWebTraffic24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthTotalWebTraffic24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthTotalWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthTotalWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthTotalWebTraffic7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthTotalWebTraffic7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgIncomingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgIncomingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgIncomingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgIncomingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgIncomingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgIncomingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgOutgoingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgOutgoingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgWebRequestSize24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgWebRequestSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgWebRequestSize24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgWebRequestSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgWebRequestSize7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgWebRequestSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgWebRequestSize7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgWebRequestSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgWebResponseSize24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgWebResponseSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgWebResponseSize24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgWebResponseSize24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthAvgWebResponseSize7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthAvgWebResponseSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthAvgWebResponseSize7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthAvgWebResponseSize7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageOKNetworkHealthSuccessHTTPRequestRatio7Days ?? throw new InvalidOperationException("Reference not set:ImageOKNetworkHealthSuccessHTTPRequestRatio7Days"), "ShowControl", (Int16)3);
            _.SETm1a0(this, _env.ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days ?? throw new InvalidOperationException("Reference not set:ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days"), "ShowControl", (Int16)3);

            // GroupBox L1-Checkliste

            _.SETm1a0(this, _env.TextBoxL1FreeSpace ?? throw new InvalidOperationException("Reference not set:TextBoxL1FreeSpace"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1OSUpToDate ?? throw new InvalidOperationException("Reference not set:TextBoxL1OSUpToDate"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Browser ?? throw new InvalidOperationException("Reference not set:TextBoxL1Browser"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Collaboration ?? throw new InvalidOperationException("Reference not set:TextBoxL1Collaboration"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Antivirus ?? throw new InvalidOperationException("Reference not set:TextBoxL1Antivirus"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Antivirus2 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Antivirus2"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Antivirus3 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Antivirus3"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Defender ?? throw new InvalidOperationException("Reference not set:TextBoxL1Defender"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1BootLogon2 ?? throw new InvalidOperationException("Reference not set:TextBoxL1BootLogon2"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1BootLogon3 ?? throw new InvalidOperationException("Reference not set:TextBoxL1BootLogon3"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1CPU24 ?? throw new InvalidOperationException("Reference not set:TextBoxL1CPU24"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1CPU7 ?? throw new InvalidOperationException("Reference not set:TextBoxL1CPU7"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Speicher24 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Speicher24"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Speicher7 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Speicher7"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Bluescreen24 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Bluescreen24"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1Bluescrren7 ?? throw new InvalidOperationException("Reference not set:TextBoxL1Bluescrren7"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1HardReset24 ?? throw new InvalidOperationException("Reference not set:TextBoxL1HardReset24"), "Text", "");
            _.SETm1a0(this, _env.TextBoxL1HardReset7 ?? throw new InvalidOperationException("Reference not set:TextBoxL1HardReset7"), "Text", "");

            // --- GroupBoxGeneralInfo

            nexthinkBaseURL = _.CONCAT(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "getNexthinkBaseURL", _.ARGS.ForceBrackets()), "query?p1=");
            nexthinkQuery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml";

            colorWarning = "#F20012";
            colorCheck = "#1B709F";

            hostname = _.VAL(_.CALLm1v5(this, _env.hlobj ?? throw new InvalidOperationException("Reference not set:hlobj"), "GetValue", "ComputerDetail.Hostname", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            if (_.IF(_.EQ(_.NullableSTR(hostname), "")))
            {
                _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", "Der Computer hat keinen Hostnamen.");
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            nexthinkURL = _.CONCAT(nexthinkBaseURL, _.UCASE(hostname), nexthinkQuery);
            nexthinkURL = "https://httpbin.org/get";

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);

            //MsgBox nexthinkURL

            //time of call
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralCallTime ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCallTime"), "Text", _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "FORMATDATETIME", _.NOW(), VBScriptConstants.vbGeneralDate)));
            });

            _.HANDLEERROR(errOn, () => {
                xmlhttp = _.OBJ(_.CREATEOBJECT("Msxml2.ServerXMLHTTP.6.0"));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v2(this, xmlhttp ?? throw new InvalidOperationException("Reference not set:xmlhttp"), "setOption", (Int16)2, (Int16)13056);
            }); //bypass certificate errors
            _.HANDLEERROR(errOn, () => {
                _.CALLm1argp(this, xmlhttp ?? throw new InvalidOperationException("Reference not set:xmlhttp"), "open", _.ARGS.Val("GET").Ref(nexthinkURL, v => { nexthinkURL = v; }).Val(false).Val(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "getNexthinkUser", _.ARGS.ForceBrackets())).Val(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "getNexthinkPassword", _.ARGS.ForceBrackets())));
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v0(this, xmlhttp ?? throw new InvalidOperationException("Reference not set:xmlhttp"), "send");
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", "Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "HelpContext")));
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
                _.SETm1a0(this, xmlDoc ?? throw new InvalidOperationException("Reference not set:xmlDoc"), "async", "false");
            });
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, xmlDoc ?? throw new InvalidOperationException("Reference not set:xmlDoc"), "load", _.CALLm1v0(this, xmlhttp ?? throw new InvalidOperationException("Reference not set:xmlhttp"), "responseXML"));
            });

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", "Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten.");
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", _.CONCAT("Error Description: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Description"), VBScriptConstants.vbLf, "Error Source: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Source"), VBScriptConstants.vbLf, "Error HelpFile: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Helpfile"), VBScriptConstants.vbLf, "Error Context: ", _.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "HelpContext")));
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
                enumerationContent = _.ENUMERABLE(_.CALLm1v1(this, xmlDoc ?? throw new InvalidOperationException("Reference not set:xmlDoc"), "SelectNodes", "//table/header/*")).GetEnumerator();
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
                    curnode = _.OBJ(_.CALLm2v1(this, xmlDoc ?? throw new InvalidOperationException("Reference not set:xmlDoc"), "documentElement", "selectSingleNode", _.CONCAT("//table/body/r/c", i)));
                });
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v2(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Add", _.CALLm1v0(this, n ?? throw new InvalidOperationException("Reference not set:n"), "Text"), _.CALLm1v0(this, curnode ?? throw new InvalidOperationException("Reference not set:curnode"), "Text"));
                });
                _.HANDLEERROR(errOn, () => {
                    i = _.ADD(i, (Int16)1);
                });
                if (enumerationContent == null)
                    break;
            }

            //Error Handling
            if (_.IF(() => _.NOTEQ(_.NullableNUM(_.CALLm1v0(this, _.ERR ?? throw new InvalidOperationException("Reference not set:ERR"), "Number")), (Int16)0), errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.model ?? throw new InvalidOperationException("Reference not set:model"), "MsgBox", "Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten.");
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return;
            }

            // from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

            //fill textboxes
            //LabelNName.Text = dict.key("name")
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralHostName ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralHostName"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralLastIP ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastIP"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "last_ip_address")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxlGeneralDeviceManufacturer ?? throw new InvalidOperationException("Reference not set:TextBoxlGeneralDeviceManufacturer"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "device_manufacturer")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralDeviceProductVersion ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralDeviceProductVersion"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "device_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralOS ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralOS"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "os_version_and_architecture")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralGroupName ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralGroupName"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "group_name")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralLastBootTime ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastBootTime"), "Text", _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "FORMATDATETIME", _.REPLACE(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "last_system_boot"), "T", " "), VBScriptConstants.vbGeneralDate)));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralLastLogon ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralLastLogon"), "Text", _.VAL(_.CALLm1v2(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "FORMATDATETIME", _.REPLACE(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "last_logon_time"), "T", " "), VBScriptConstants.vbGeneralDate)));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralDeviceType ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralDeviceType"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "last_logged_on_user")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralBIOSSerialNumber ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralBIOSSerialNumber"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "bios_serial_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralCPUModel ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCPUModel"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "cpu_model")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralNumberOfCPUs ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfCPUs"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "number_of_cpus")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralNumberOfLogProcs ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfLogProcs"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "logical_cpu_number")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralNumberOfCores ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfCores"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "number_of_cores")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralCPUFreq ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralCPUFreq"), "Text", _.CONCAT(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "cpu_frequency"), " MHz"));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralTotalRAM ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralTotalRAM"), "Text", _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ConvertSize", _.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "total_ram"))));
            });

            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralNumberOfGraphCards ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralNumberOfGraphCards"), "Text", _.VAL(_.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "number_of_graphical_cards")));
            });
            _.HANDLEERROR(errOn, () => {
                _.SETm1a0(this, _env.TextBoxGeneralGraphCardRAM ?? throw new InvalidOperationException("Reference not set:TextBoxGeneralGraphCardRAM"), "Text", _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "ConvertSize", _.CALLm1v1(this, dict ?? throw new InvalidOperationException("Reference not set:dict"), "Item", "graphical_card_ram"))));
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
