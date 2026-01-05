using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.Tests.Application.Controls;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class DialogGui : TestBase
    {
        [TestMethod]
        public void QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click()// => TestDialogGui();
        //private void TestDialogGui()
        {
            var dialog = new DialogBuilder()
                .AddTextControl("TextBoxWebsite")
                .BuildDialog();
            var externalReferences = dialog.Controls;
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, externalReferences);
            DoDialogGui(externalReferences);
        }

        [TestMethod]
        public void CT35_LogChecklist_Dialog_388_OnSave() // 35:DFSnDLNeu
        {
            var dialog = new DialogBuilder()
                .AddTextControl("TextBoxChecklist1URL")
                .AddTextControl("TextBoxChecklist2URL")
                .AddTextControl("TextBoxChecklist3URL")
                .AddTextControl("TextBoxChecklist4URL")
                .AddTextControl("TextBoxChecklist5URL")
                .AddTextControl("TextBoxChecklist6URL")
                .AddTextControl("TextBoxChecklist7URL")
                .AddTextControl("TextBoxChecklist8URL")
                .AddTextControl("TextBoxChecklist9URL")
                .AddTextControl("TextBoxChecklist10URL")
                .BuildDialog();

            var externalReferences = dialog.Controls;
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, externalReferences);
            DoDialogGui(externalReferences);

        }

        //[TestMethod]
        public void CT125_ClientComputer_Dialog_349_ButtonGeneralInfo_Click()
        {
            var dialog = new DialogBuilder()
                    .AddTabControl("TabPageGeneralInfo")
                    .AddTextControl("TextBoxChecklist2URL")

                    .AddTabControl("TabPageSoftwareOSHealth")
                    .AddTabControl("TabPageSecurityCompliance")
                    .AddTabControl("TabPageTechnicalInfo")
                    .AddTabControl("TabPageNetworkHealth")
                    .AddTabControl("TabPageL1Checklist")
                    .AddGroupBox("GroupBoxGeneralInfo")
                    .AddGroupBox("GroupBoxTechnicalInfo")
                    .AddGroupBox("GroupBoxSoftwareOSHealth")
                    .AddGroupBox("GroupBoxSecurityCompliance")
                    .AddGroupBox("GroupBoxNetworkHealth")
                    .AddGroupBox("GroupBoxL1Checklist")
                    .AddButton("ButtonGeneralInfo")
                    .AddButton("ButtonTechnicalInfo")
                    .AddButton("ButtonSWHealth")
                    .AddButton("ButtonSecurityCompliance")
                    .AddButton("ButtonNetworkHealth")
                    .AddButton("ButtonL1Checklist")
                    .AddTextControl("TabControlNexthink")
                    .AddTextControl("TextBoxGeneralCallTime")
                    .AddTextControl("TextBoxGeneralHostName")
                    .AddTextControl("TextBoxlGeneralDeviceManufacturer")
                    .AddTextControl("TextBoxGeneralDeviceProductVersion") /* (line 32) */
                    .AddTextControl("TextBoxGeneralLastIP") /* (line 33) */
                    .AddTextControl("TextBoxGeneralGroupName") /* (line 34) */
                    .AddTextControl("TextBoxGeneralOS") /* (line 35) */
                    .AddTextControl("TextBoxGeneralLastBootTime") /* (line 36) */
                    .AddTextControl("TextBoxGeneralLastLogon") /* (line 37) */
                    .AddTextControl("TextBoxGeneralDeviceType") /* (line 38) */
                    .AddTextControl("TextBoxGeneralBIOSSerialNumber") /* (line 39) */
                    .AddTextControl("TextBoxGeneralCPUModel") /* (line 40) */
                    .AddTextControl("TextBoxGeneralNumberOfCPUs") /* (line 41) */
                    .AddTextControl("TextBoxGeneralNumberOfLogProcs") /* (line 42) */
                    .AddTextControl("TextBoxGeneralNumberOfCores") /* (line 43) */
                    .AddTextControl("TextBoxGeneralCPUFreq") /* (line 44) */
                    .AddTextControl("TextBoxGeneralTotalRAM") /* (line 45) */
                    .AddTextControl("TextBoxGeneralNumberOfGraphCards") /* (line 46) */
                    .AddTextControl("TextBoxTechnicalInfoTotalDriveCapNow") /* (line 50) */
                    .AddTextControl("TextBoxTechnicalInfoTotalFreeSpaceNow") /* (line 51) */
                    .AddTextControl("TextBoxTechnicalInfoTotalDriveUsageNow") /* (line 52) */
                    .AddTextControl("TextBoxTechnicalInfoSystemDriveCapNow") /* (line 53) */
                    .AddTextControl("TextBoxTechnicalInfoSystemDriveFreeSpaceNow") /* (line 54) */
                    .AddTextControl("TextBoxTechnicalInfoHighCPUTimeNow") /* (line 55) */
                    .AddTextControl("TextBoxTechnicalInfoHighMemoryTimeNow") /* (line 56) */
                    .AddTextControl("TextBoxTechnicalInfoHighIOTimeNow") /* (line 57) */
                    .AddTextControl("TextBoxTechnicalInfoTotalDriveCap7Days") /* (line 58) */
                    .AddTextControl("TextBoxTechnicalInfoTotalFreeSpace7Days") /* (line 59) */
                    .AddTextControl("TextBoxTechnicalInfoTotalDriveUsage7Days") /* (line 60) */
                    .AddTextControl("TextBoxTechnicalInfoSystemDriveCap7Days") /* (line 61) */
                    .AddTextControl("TextBoxTechnicalInfoSystemDriveFreeSpace7Days")
                    .AddTextControl("TextBoxTechnicalInfoHighCPUTime7Days")
                    .AddTextControl("TextBoxTechnicalInfoHighMemoryTime7Days")
                    .AddTextControl("TextBoxTechnicalInfoHighIOTime7Days")
                    .AddImageControl("ImageNOKTechnicalInfoTotalFreeSpaceNow") /* (line 67) */
                    .AddImageControl("ImageOKTechnicalInfoTotalFreeSpaceNow") /* (line 68) */
                    .AddImageControl("ImageNOKTechnicalInfoTotalDriveUsageNow") /* (line 69) */
                    .AddImageControl("ImageOKTechnicalInfoTotalDriveUsageNow") /* (line 70) */
                    .AddImageControl("ImageNOKTechnicalInfoSystemDriveCapNow") /* (line 71) */
                    .AddImageControl("ImageOKTechnicalInfoSystemDriveCapNow") /* (line 72) */
                    .AddImageControl("ImageNOKTechnicalInfoSystemDriveFreeSpaceNow") /* (line 73) */
                    .AddImageControl("ImageOKTechnicalInfoSystemDriveFreeSpaceNow") /* (line 74) */
                    .AddImageControl("ImageNOKTechnicalInfoHighCPUTimeNow") /* (line 75) */
                    .AddImageControl("ImageOKTechnicalInfoHighCPUTimeNow") /* (line 76) */
                    .AddImageControl("ImageNOKTechnicalInfoHighMemoryTimeNow") /* (line 77) */
                    .AddImageControl("ImageOKTechnicalInfoHighMemoryTimeNow") /* (line 78) */
                    .AddImageControl("ImageNOKTechnicalInfoHighIOTimeNow") /* (line 79) */
                    .AddImageControl("ImageOKTechnicalInfoHighIOTimeNow") /* (line 80) */
                    .AddImageControl("ImageOKTechnicalInfoTotalFreeSpace7Days") /* (line 82) */
                    .AddImageControl("ImageNOKTechnicalInfoTotalDriveUsage7Days") /* (line 83) */
                    .AddImageControl("ImageOKTechnicalInfoTotalDriveUsage7Days") /* (line 84) */
                    .AddImageControl("ImageNOKTechnicalInfoSystemDriveCap7Days") /* (line 85) */
                    .AddImageControl("ImageOKTechnicalInfoSystemDriveCap7Days") /* (line 86) */
                    .AddImageControl("ImageNOKTechnicalInfoSystemDriveFreeSpace7Days") /* (line 87) */
                    .AddImageControl("ImageOKTechnicalInfoSystemDriveFreeSpace7Days") /* (line 88) */
                    .AddImageControl("ImageNOKTechnicalInfoHighCPUTime7Days") /* (line 89) */
                    .AddImageControl("ImageOKTechnicalInfoHighCPUTime7Days") /* (line 90) */
                    .AddImageControl("ImageNOKTechnicalInfoHighMemoryTime7Days") /* (line 91) */
                    .AddImageControl("ImageOKTechnicalInfoHighMemoryTime7Days") /* (line 92) */
                    .AddImageControl("ImageNOKTechnicalInfoHighIOTime7Days") /* (line 93) */
                    .AddImageControl("ImageOKTechnicalInfoHighIOTime7Days") /* (line 94) */
                    .AddTextControl("TextBoxSoftwareOSHealthOSVersionArchitecture") /* (line 98) */
                    .AddTextControl("TextBoxSoftwareOSHealthOSName") /* (line 99) */
                    .AddTextControl("TextBoxSoftwareOSHealthWMIStatus") /* (line 100) */
                    .AddTextControl("TextBoxSoftwareOSHealthLastSystemUpdate") /* (line 101) */
                    .AddTextControl("TextBoxSoftwareOSHealthWindowsUpdateStatus") /* (line 102) */
                    .AddTextControl("TextBoxSoftwareOSHealthNumberOfApps") /* (line 103) */
                    .AddTextControl("TextBoxSoftwareOSHealthNumberOfExes") /* (line 104) */
                    .AddTextControl("TextBoxSoftwareOSHealthNumberOfBins") /* (line 105) */
                    .AddTextControl("TextBoxSoftwareOSHealthOSEndOfSupport") /* (line 106) */
                    .AddTextControl("TextBoxSoftwareOSHealthOSIE11Support") /* (line 107) */
                    .AddTextControl("TextBoxSoftwareOSHealthWin10Ready") /* (line 108) */
                    .AddTextControl("TextBoxSoftwareOSHealthOSComplience") /* (line 109) */
                    .AddTextControl("TextBoxSecurityComplianceInetSecuritySettings") /* (line 113) */
                    .AddTextControl("TextBoxSecurityComplianceUserAccountStatus") /* (line 114) */
                    .AddTextControl("TextBoxSecurityComplianceAntivirusName") /* (line 115) */
                    .AddTextControl("TextBoxSecurityComplianceAntivirusRTP") /* (line 116) */
                    .AddTextControl("TextBoxSecurityComplianceAntivirusUpToDate") /* (line 117) */
                    .AddTextControl("TextBoxSecurityComplianceAntivirusNumber") /* (line 118) */
                    .AddTextControl("TextBoxSecurityComplianceAntivirusAll") /* (line 119) */
                    .AddTextControl("TextBoxSecurityComplianceAntispywareName") /* (line 120) */
                    .AddTextControl("TextBoxSecurityComplianceAntispywareRTP") /* (line 121) */
                    .AddTextControl("TextBoxSecurityComplianceAntispywareUpToDate") /* (line 122) */
                    .AddTextControl("TextBoxSecurityComplianceAntispywareNumber") /* (line 123) */
                    .AddTextControl("TextBoxSecurityComplianceAntispywareAll") /* (line 124) */
                    .AddTextControl("TextBoxSecurityComplianceFirewallName") /* (line 125) */
                    .AddTextControl("TextBoxSecurityComplianceFirewallRTP") /* (line 126) */
                    .AddTextControl("TextBoxSecurityComplianceFirewallNumber") /* (line 127) */
                    .AddTextControl("TextBoxSecurityComplianceFirewallAll") /* (line 128) */
                    .AddTextControl("TextBoxNetworkHealthIncomingNetTaffic24Hours") /* (line 131) */
                    .AddTextControl("TextBoxNetworkHealthOutgoingNetTaffic24Hours") /* (line 132) */
                    .AddTextControl("TextBoxNetworkHealthTotalNetTaffic24Hours") /* (line 133) */
                    .AddTextControl("TextBoxNetworkHealthSuccessNetConnectionRatio24Hours") /* (line 134) */
                    .AddTextControl("TextBoxNetworkHealthNetAvailLevel24Hours") /* (line 135) */
                    .AddTextControl("TextBoxNetworkHealthAvgIncomingNetBitrate24Hours") /* (line 136) */
                    .AddTextControl("TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours") /* (line 137) */
                    .AddTextControl("TextBoxNetworkHealthAvgNetResponseTime24Hours") /* (line 138) */
                    .AddTextControl("TextBoxNetworkHealthIncomingWebTraffic24Hours") /* (line 139) */
                    .AddTextControl("TextBoxNetworkHealthOutgoingWebTraffic24Hours") /* (line 140) */
                    .AddTextControl("TextBoxNetworkHealthTotalWebTraffic24Hours") /* (line 141) */
                    .AddTextControl("TextBoxNetworkHealthAvgIncomingWebBitrate24Hours") /* (line 142) */
                    .AddTextControl("TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours") /* (line 143) */
                    .AddTextControl("TextBoxNetworkHealthAvgWebRequestSize24Hours") /* (line 144) */
                    .AddTextControl("TextBoxNetworkHealthAvgWebResponseSize24Hours") /* (line 145) */
                    .AddTextControl("TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours") /* (line 146) */
                    .AddTextControl("TextBoxNetworkHealthIncomingNetTaffic7Days") /* (line 148) */
                    .AddTextControl("TextBoxNetworkHealthOutgoingNetTaffic7Days") /* (line 149) */
                    .AddTextControl("TextBoxNetworkHealthTotalNetTaffic7Days") /* (line 150) */
                    .AddTextControl("TextBoxNetworkHealthSuccessNetConnectionRatio7Days") /* (line 151) */
                    .AddTextControl("TextBoxNetworkHealthNetAvailLevel7Days") /* (line 152) */
                    .AddTextControl("TextBoxNetworkHealthAvgIncomingNetBitrate7Days") /* (line 153) */
                    .AddTextControl("TextBoxNetworkHealthAvgOutgoingNetBitrate7Days") /* (line 154) */
                    .AddTextControl("TextBoxNetworkHealthAvgNetResponseTime7Days") /* (line 155) */
                    .AddTextControl("TextBoxNetworkHealthIncomingWebTraffic7Days") /* (line 156) */
                    .AddTextControl("TextBoxNetworkHealthOutgoingWebTraffic7Days")

                    .AddTextControl("TextBoxNetworkHealthTotalWebTraffic7Days")
                    .AddTextControl("TextBoxNetworkHealthAvgIncomingWebBitrate7Days")
                    .AddTextControl("TextBoxNetworkHealthAvgOutgoingWebBitrate7Days")
                    .AddTextControl("TextBoxNetworkHealthAvgWebRequestSize7Days")
                    .AddTextControl("TextBoxNetworkHealthAvgWebResponseSize7Days")
                    .AddTextControl("TextBoxNetworkHealthSuccessHTTPRequestRatio7Days")

.AddImageControl("ImageOKNetworkHealthIncomingNetTaffic24Hours")
.AddImageControl("ImageNOKNetworkHealthIncomingNetTaffic24Hours")
.AddImageControl("ImageOKNetworkHealthIncomingNetTaffic7Days")
.AddImageControl("ImageNOKNetworkHealthIncomingNetTaffic7Days")
.AddImageControl("ImageOKNetworkHealthOutgoingNetTaffic24Hours")
.AddImageControl("ImageNOKNetworkHealthOutgoingNetTaffic24Hours")
.AddImageControl("ImageOKNetworkHealthOutgoingNetTaffic7Days")
.AddImageControl("ImageNOKNetworkHealthOutgoingNetTaffic7Days")
.AddImageControl("ImageOKNetworkHealthTotalNetTaffic24Hours")
.AddImageControl("ImageNOKNetworkHealthTotalNetTaffic24Hours")
.AddImageControl("ImageOKNetworkHealthTotalNetTaffic7Days")
.AddImageControl("ImageNOKNetworkHealthTotalNetTaffic7Days")
.AddImageControl("ImageOKNetworkHealthSuccessNetConnectionRatio24Hours")
.AddImageControl("ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours")
.AddImageControl("ImageOKNetworkHealthSuccessNetConnectionRatio7Days")
.AddImageControl("ImageNOKNetworkHealthSuccessNetConnectionRatio7Days")
.AddImageControl("ImageOKNetworkHealthNetAvailLevel24Hours")
.AddImageControl("ImageNOKNetworkHealthNetAvailLevel24Hours")
.AddImageControl("ImageOKNetworkHealthNetAvailLevel7Days")
.AddImageControl("ImageNOKNetworkHealthNetAvailLevel7Days")
.AddImageControl("ImageOKNetworkHealthAvgIncomingNetBitrate24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours")
.AddImageControl("ImageOKNetworkHealthAvgIncomingNetBitrate7Days")
.AddImageControl("ImageNOKNetworkHealthAvgIncomingNetBitrate7Days")
.AddImageControl("ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours")
.AddImageControl("ImageOKNetworkHealthAvgOutgoingNetBitrate7Days")
.AddImageControl("ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days")
.AddImageControl("ImageOKNetworkHealthAvgNetResponseTime24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgNetResponseTime24Hours")
.AddImageControl("ImageOKNetworkHealthAvgNetResponseTime7Days")
.AddImageControl("ImageNOKNetworkHealthAvgNetResponseTime7Days")
.AddImageControl("ImageOKNetworkHealthIncomingWebTraffic24Hours")
.AddImageControl("ImageNOKNetworkHealthIncomingWebTraffic24Hours")
.AddImageControl("ImageOKNetworkHealthIncomingWebTraffic7Days")
.AddImageControl("ImageNOKNetworkHealthIncomingWebTraffic7Days")
.AddImageControl("ImageOKNetworkHealthOutgoingWebTraffic24Hours")
.AddImageControl("ImageNOKNetworkHealthOutgoingWebTraffic24Hours")
.AddImageControl("ImageOKNetworkHealthOutgoingWebTraffic7Days")
.AddImageControl("ImageNOKNetworkHealthOutgoingWebTraffic7Days")
.AddImageControl("ImageOKNetworkHealthTotalWebTraffic24Hours")
.AddImageControl("ImageOKNetworkHealthTotalWebTraffic7Days")
.AddImageControl("ImageNOKNetworkHealthTotalWebTraffic7Days")
.AddImageControl("ImageOKNetworkHealthAvgIncomingWebBitrate24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours")
.AddImageControl("ImageOKNetworkHealthAvgIncomingWebBitrate7Days")
.AddImageControl("ImageNOKNetworkHealthAvgIncomingWebBitrate7Days")
.AddImageControl("ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours")
.AddImageControl("ImageOKNetworkHealthAvgOutgoingWebBitrate7Days")
.AddImageControl("ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days")
.AddImageControl("ImageOKNetworkHealthAvgWebRequestSize24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgWebRequestSize24Hours")
.AddImageControl("ImageOKNetworkHealthAvgWebRequestSize7Days")
.AddImageControl("ImageNOKNetworkHealthAvgWebRequestSize7Days")
.AddImageControl("ImageOKNetworkHealthAvgWebResponseSize24Hours")
.AddImageControl("ImageNOKNetworkHealthAvgWebResponseSize24Hours")
.AddImageControl("ImageOKNetworkHealthAvgWebResponseSize7Days")
.AddImageControl("ImageNOKNetworkHealthAvgWebResponseSize7Days")
.AddImageControl("ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours")
.AddImageControl("ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours")
.AddImageControl("ImageOKNetworkHealthSuccessHTTPRequestRatio7Days")
.AddImageControl("ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days")

// GroupBox L1-Checkliste

.AddTextControl("TextBoxL1FreeSpace")
.AddTextControl("TextBoxL1OSUpToDate")
.AddTextControl("TextBoxL1Browser")
.AddTextControl("TextBoxL1Collaboration")
.AddTextControl("TextBoxL1Antivirus")
.AddTextControl("TextBoxL1Antivirus2")
.AddTextControl("TextBoxL1Antivirus3")
.AddTextControl("TextBoxL1Defender")
.AddTextControl("TextBoxL1BootLogon2")
.AddTextControl("TextBoxL1BootLogon3")
.AddTextControl("TextBoxL1CPU24")
.AddTextControl("TextBoxL1CPU7")
.AddTextControl("TextBoxL1Speicher24")
.AddTextControl("TextBoxL1Speicher7")
.AddTextControl("TextBoxL1Bluescreen24")
.AddTextControl("TextBoxL1Bluescrren7")
.AddTextControl("TextBoxL1HardReset24")
.AddTextControl("TextBoxL1HardReset7")

.AddLabelControl("TextBoxGeneralGraphCardRAM")

                    .BuildDialog();

            var hlobj = new HLObjectInstance()
                .RegisterValueKey<string>("ComputerDetail.Hostname", 0, 0, "hst-X_1")
                ;
            var model = new DialogGuidModel();

            var externalReferences = dialog.Controls;
            externalReferences.Add("hlobj", hlobj);
            externalReferences.Add("model", model);
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.Controls);
            DoDialogGui(externalReferences);
        }

        private void DoDialogGui(Dictionary<string, object> externalReferences)
        {
            CncIn.ExecuteTranslatedProgram(RuntimeLogger, TestCulture, TestContext.TestName, externalReferences);
        }
    }

    [ComVisible(true)]
    internal sealed class DialogGuidModel
    {
        public DialogGuidModel()
        {
        }

        public void MsgBox(string message)
        {
            Console.WriteLine($"MsgBox('{message}')");
        }
    }
}
