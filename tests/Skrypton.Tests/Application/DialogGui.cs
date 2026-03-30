using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application.Controls;
using Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests;
using Skrypton.Tests.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations.FileSystemSupport;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class DialogGui : TestBase
    {
        [TestMethod]
        public void QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click()// => TestDialogGui();
        {
            var model = new DialogGuidModel(TestCulture);
            var dialog = new DialogBuilder(CreateTestHostServices(), model)
                .AddTextControl("TextBoxWebsite")
                .BuildDialog();
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(this, rsp.TranslatedCsCode, dialog, gr => { });
        }

        [TestMethod]
        public void CT35_LogChecklist_Dialog_388_OnSave() // 35:DFSnDLNeu  id = select id, dbname from _databasestats order by dbname asc -- [hlsysdialog]
        {
            var model = new DialogGuidModel(TestCulture);
            var dialog = new DialogBuilder(CreateTestHostServices(), model)
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

            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences, suppressions: ["SKY101", "SKY105", "SKY106"]);
            DoDialogGui(this, rsp.TranslatedCsCode, dialog, gr => { });

        }

        void XX()
        {
            /*
Dim xmlhttp
Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")

' Optional: bypass certificate errors (not needed for HTTP, but included for parity)
xmlhttp.setOption 2, 13056

Dim testUrl
testUrl = "https://httpbin.org/get"

' Public URL, no authentication required
xmlhttp.open "GET", testUrl, False, "", ""

xmlhttp.send

WScript.Echo "Status: " & xmlhttp.status
WScript.Echo "Response:"
WScript.Echo xmlhttp.responseText
             */
        }

        static void TestCS()
        {
            // Create COM object
#pragma warning disable CA1416 // Validate platform compatibility
            Type t = typeof(MyServerXMLHTTP60);//Type.GetTypeFromProgID("Msxml2.ServerXMLHTTP.6.0", true);
#pragma warning restore CA1416 // Validate platform compatibility
            dynamic xmlhttp = Activator.CreateInstance(t);

            // 2 = SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
            // 13056 = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
            xmlhttp.setOption(2, 13056);

            string url = "https://httpbin.org/get";

            // Open synchronous GET request, no authentication needed
            xmlhttp.open("GET", url, false, "myusr2", "mypwd2");

            // Send request
            xmlhttp.send();

            Console.WriteLine("Status: " + xmlhttp.status);
            Console.WriteLine("Response:");
            Console.WriteLine("responseText:" + xmlhttp.responseText);
        }

        static void TestCS_IDispatch()
        {
            // Create COM object
#pragma warning disable CA1416 // Validate platform compatibility
            Type t = typeof(MyServerXMLHTTP60);// Type.GetTypeFromProgID("Msxml2.ServerXMLHTTP.6.0", true);
#pragma warning restore CA1416 // Validate platform compatibility
            object xmlhttp = Activator.CreateInstance(t);

            // xmlhttp.setOption(2, 13056)
            ComInvoke.Call(xmlhttp, "setOption", [2, 13056]);

            string url = "https://httpbin.org/get";

            // xmlhttp.open("GET", url, false, "myusr2", "mypwd2")
            ComInvoke.Call(xmlhttp, "open", ["GET", url, false, "myusr2", "mypwd2"]);

            // xmlhttp.send()
            ComInvoke.Call(xmlhttp, "send", []);

            // Read status
            object status = xmlhttp.GetType().InvokeMember(
                "status",
                BindingFlags.GetProperty,
                null,
                xmlhttp,
                null
            );

            // Read responseText
            object responseText = xmlhttp.GetType().InvokeMember(
                "responseText",
                BindingFlags.GetProperty,
                null,
                xmlhttp,
                null
            );

            Console.WriteLine("Status: " + status);
            Console.WriteLine("Response:");
            Console.WriteLine(responseText);
        }
        static void Test_IDispatch_Invoke()
        {
            //works only on windows
            //if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            object comObj = DefaultRuntimeFunctionalityProvider.TestCreateComObjectTest("Msxml2.ServerXMLHTTP.6.0", typeof(MyServerXMLHTTP60));
            // Create COM object
            //#pragma warning disable CA1416 // Validate platform compatibility
            //Type t = typeof(MyServerXMLHTTP60);// Type.GetTypeFromProgID("Msxml2.ServerXMLHTTP.6.0", true);
            //#pragma warning restore CA1416 // Validate platform compatibility
            //object comObj = Activator.CreateInstance(t);

            // Cast to IDispatch
            //var disp = (IDispatchAccess.IDispatch)comObj;
            var disp = (IReflect)comObj;

            // --- setOption(2, 13056) ---
            InvokeMethod(disp, "setOption", 2, 13056);
            //IDispatchAccess.CallMethod(disp, "setOption", 2, 13056);

            string url = "https://httpbin.org/get";

            // --- open("GET", url, false, "myusr2", "mypwd2") ---
            InvokeMethod(disp, "open", "GET", url, false, "myusr2", "mypwd2");
            //IDispatchAccess.CallMethodU(disp, "open", "GET", url, false, "myusr2", "mypwd2");

            // --- send() ---
            InvokeMethod(disp, "send");
            //IDispatchAccess.CallMethodU(disp, "send");

            // --- status property ---
            object status = InvokePropertyGet(disp, "status");
            //object status = IDispatchAccess.GetProperty(disp, "status");

            // --- responseText property ---
            object responseText = InvokePropertyGet(disp, "responseText");
            //object responseText = IDispatchAccess.GetProperty(disp, "responseText");

            Console.WriteLine("Status: " + status);
            Console.WriteLine("Response:");
            Console.WriteLine(responseText);
        }

        private static void Test_Xml()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(MyServerXMLHTTP60.SampleResponseXml);

            var dict = new Dictionary<string, string>();

            XmlNodeList headers = xmlDoc.SelectNodes("//table/header/*");

            int i = 0;
            foreach (XmlNode header in headers)
            {
                XmlNode valueNode = xmlDoc.SelectSingleNode($"//table/body/r/c{i}");
                dict[header.InnerText] = valueNode?.InnerText ?? string.Empty;
                i++;
            }
        }

        private static void Test_Xml_IDispatch()
        {
#pragma warning disable CA1416 // Validate platform compatibility
            Type domType = typeof(MyMsxml2DOMDocument); // Type.GetTypeFromProgID("Msxml2.DOMDocument"); // IXMLDOMDocument

#pragma warning restore CA1416 // Validate platform compatibility
            object xmlDoc = Activator.CreateInstance(domType);

            domType.InvokeMember("setProperty", BindingFlags.InvokeMethod, null, xmlDoc, ["SelectionLanguage", "XPath"]);
            // xmlDoc.async = false
            domType.InvokeMember("async", BindingFlags.SetProperty, null, xmlDoc, [false]);

            // !!!! VBScript silently sets:             xmlDoc.preserveWhiteSpace = False
            // The behavior comes from three different layers that were never fully documented together:
            // * 1/3) VBScript’s automatic type coercion rules. Key point: VBScript will coerce strings to booleans, numbers, objects, etc. when calling COM methods. https://learn.microsoft.com/en-us/previous-versions//d1wf56tt(v=vs.85)
            // * 2/3) MSXML’s COM overloading rules. VBScript uses IDispatch::Invoke with very permissive rules. https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke
            //  -> unwraps COM objects automatically
            //  -> chooses the correct overloaded COM method
            //  -> retries calls with different type coercions
            //  -> suppresses many COM errors
            // * 3/3) The IDispatch binder inside Windows Script Host
            domType.InvokeMember("preserveWhiteSpace", BindingFlags.SetProperty, null, xmlDoc, [false]);

            //object responseXML = MyServerXMLHTTP60.SampleResponseXml;
            object responseXML = Encoding.UTF8.GetBytes(MyServerXMLHTTP60.SampleResponseXml);
            //xmlSource[in] An indicator of the source XML to parse. This may be an URL(String / BSTR), a Request object(in an ASP page), an IStream, SAFEARRAY of bytes(VT_ARRAY| VT_UI1), a DOMDocument object, or any object that supports IStream, ISequentialStream, or IPersistStream.See Remarks for more information.
            domType.InvokeMember("load", BindingFlags.InvokeMethod, null, xmlDoc, [responseXML]);

            // Select header nodes
            object headerNodes = domType.InvokeMember(
                "selectNodes",
                BindingFlags.InvokeMethod,
                null,
                xmlDoc,
                ["//table/header/*"]
            );
            // Iterate header nodes
            int length = (int)headerNodes.GetType().InvokeMember("length", BindingFlags.GetProperty, null, headerNodes, null);
            if (length == 0)
                throw new NotImplementedException(); // loadXML vs load

            // Get documentElement
            object docElem = domType.InvokeMember("documentElement", BindingFlags.GetProperty, null, xmlDoc, null);

            int i = 0;

            for (int idx = 0; idx < length; idx++)
            {
                // n = headerNodes.item(idx)
                object n = headerNodes.GetType().InvokeMember("item", BindingFlags.InvokeMethod, null, headerNodes, [idx]);

                // curnode = xmlDoc.documentElement.selectSingleNode("//table/body/r/c" & i)
                object curnode = docElem.GetType().InvokeMember(
                    "selectSingleNode",
                    BindingFlags.InvokeMethod,
                    null,
                    docElem,
                    ["//table/body/r/c" + i]
                );

                // dict.Add n.text, curnode.text

                var key = n.GetType().InvokeMember("text", BindingFlags.GetProperty, null, n, null);
                var value = curnode.GetType().InvokeMember("text", BindingFlags.GetProperty, null, curnode, null);
                Console.WriteLine($"{i} {key}:{value}");

                i++;
            }
        }

        [TestMethod]
        public void CT125_ClientComputer_Dialog_349_ButtonGeneralInfo_Click()
        {
            Test_Xml_IDispatch();
            Test_Xml();
            TestCS();
            TestCS_IDispatch();
            //Test_IDispatch_Invoke(); //works only on windows

            var hlobj = new HLObjectInstance()
                    .RegisterValueKey<string>("ComputerDetail.Hostname", 0, 0, "hst-X_1")
                ;
            var model = new DialogGuidModel(TestCulture);

            var dialog = new DialogBuilder(CreateTestHostServices(), model).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
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


            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences, suppressions: ["SKY102", "SKY103", "SKY105", "SKY106"]);
            DoDialogGui(this, rsp.TranslatedCsCode, dialog, gr => { });
        }

        private void DoDialogGui(DialogBase dialog, string[] suppressions, Action<GlobalReferencesBase> dialogHandler)
        {
            string translated_cs = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension, isOptional: true);
            if (translated_cs == null)
            {
                Console.WriteLine("translating...");
                string scriptContent = dialog.CompleteScriptCode();
                translated_cs = DefaultCSharpTranslation.GetTranslatedProgramCode(TestCulture, scriptContent, dialog.ExternalReferences.Keys.ToArray(), [], suppressions);
            }

            DoDialogGui(this, translated_cs, dialog, dialogHandler);
        }
        private static void DoDialogGui(TestBaseX tst, string translated_cs, DialogBase dialog, Action<GlobalReferencesBase> dialogHandler)
        {
            CncIn.ExecuteTranslatedProgram(tst, translated_cs, dialog.HostServices, dialog.ExternalReferences, dialogHandler);
        }

        static object InvokePropertyGet(IDispatchAccess.IDispatch disp, string name)
        {
            return IDispatchAccess.GetProperty(disp, name);
        }

        internal const BindingFlags BindingFlagsVBScript = BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase;
        static object InvokePropertyGet(IReflect disp, string name)
        {
            return disp.GetProperty(name, BindingFlagsVBScript);
        }
        static object InvokeMethod(IReflect disp, string name, params object[] args)
        {
            var method = disp.GetMethod(name, BindingFlagsVBScript);
            return method.Invoke(disp, args);
        }

        // Resolve DISPID and call a method
        static object InvokeMethod(IDispatchAccess.IDispatch disp, string name, params object[] args)
        {
            //IReflect reflect =
            int dispid = IDispatchAccess.GetDispId(disp, name);

            // COM expects arguments in reverse order
            Array.Reverse(args);

            var dispParams = new DISPPARAMS
            {
                cArgs = args.Length,
                rgvarg = MarshalArgs(args),
                cNamedArgs = 0,
                rgdispidNamedArgs = IntPtr.Zero
            };
            Guid GuidEmpty = Guid.Empty;
            var excepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            object result;
            var hrRet = ((IDispatchAccess.IDispatch)disp).Invoke(
                dispid,
                ref GuidEmpty,
                0,
                (ushort)INVOKEKIND.INVOKE_FUNC,
                ref dispParams,
                out result,
                ref excepInfo,
                out var argErr
            );
            if (hrRet != 0)
            {
                throw new InvalidOperationException($"HR:{hrRet}");
            }

            IDispatchAccess.FreeDISPPARAMS(dispParams);
            return result;
        }

        // Marshal arguments into unmanaged VARIANT array
        static IntPtr MarshalArgs(object[] args)
        {
            int size = Marshal.SizeOf<VARIANT>();
            IntPtr ptr = Marshal.AllocCoTaskMem(size * args.Length);

            for (int i = 0; i < args.Length; i++)
            {
                IntPtr p = IntPtr.Add(ptr, i * size);
                VARIANT variant = VARIANT.FromObject(args[i]);
                Marshal.StructureToPtr(variant, p, false);
            }

            return ptr;
        }

        [TestMethod]
        public void CT130_ClientComputer_Dialog_567_Button1_Click() // select content_uncompressed from hlsysdialog where dbname = '_CustomerTest_Tamedia' and dialogid = 567; select * from hlsysdialogglobalscript where dbname = '_CustomerTest_Tamedia';
        {
            var hlobj = new HLObjectInstance()
                    .RegisterValueKey<string>("vRealize.LansweeperURL", 0, 0, "hst-X_1")
                ;
            var model = new DialogGuidModel(TestCulture);

            IHostProcessControlHostService processControlHostService = CreateTestProcessControlHostService();

            var dialog = new DialogBuilder(CreateTestHostServices(r => r.RegisterHostService<IHostProcessControlHostService>(() => processControlHostService)), model).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
                .AddButton("Button1_Click")
                .BuildDialog();

            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences, suppressions: ["SKY102", "SKY105", "SKY106", "SKY108"]);
            DoDialogGui(this, rsp.TranslatedCsCode, dialog, (GlobalReferencesBase gr) =>
            {
                var mis = gr.GetType().GetMethods().OrderBy(x => x.Name).ToArray();
                foreach (var mi in mis)
                {
                    if (mi.DeclaringType == gr.GetType())
                    {
                        Console.WriteLine($"Method: {mi.Name}");
                    }
                }

                ScriptControlClass.RunProcedure(gr, "Button1_click", parameters: []);
            });
        }

        [TestMethod]
        public void CT74_ClientComputer_Dialog_2_ButtonShowWebsite_Click() // select * from hlsysdialog where dbname = '_CustomerTest_Mainova' and dialogid = 2; select * from hlsysdialogglobalscript where dbname = '_CustomerTest_Mainova';
        {
            var hlobj = new HLObjectInstance().InitializeObjectInstance(isNew: true, TestCulture)
                    .RegisterValueKey<string>("PersonBilling.CostCenter_CA", 0, 0, "hst-X_1")
                    .RegisterValueKey<string>("PersonInformation.SBCode", 0, 0, "hst-X_1")
                    .RegisterValueKey<string>("PersonGeneral.Name", 0, 0, "Kuku-Muku")
                    .RegisterValueKey<string>("PersonGeneral.Group", 0, 0, "g-x1")
                    .RegisterValueKey<string>("PersonGeneral.PersonalID", 0, 0, "prsnid-x1")
                ;
            var model = new DialogGuidModel(TestCulture);

            IHostDatabaseConnectionFactoryHostService databaseConnectionFactoryHostService = CreateTestDatabaseConnectionFactoryHostService();

            var dialog = new DialogBuilder(CreateTestHostServices(r => r.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => databaseConnectionFactoryHostService)), model).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
                .AddButton("ButtonShowWebsite_Click")
                .BuildDialog();

            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences, suppressions: ["SKY106"]);
            DoDialogGui(this, rsp.TranslatedCsCode, dialog, (GlobalReferencesBase gr) =>
            {
                var mis = gr.GetType().GetMethods().OrderBy(x => x.Name).ToArray();
                foreach (var mi in mis)
                {
                    if (mi.DeclaringType == gr.GetType())
                    {
                        Console.WriteLine($"Method: {mi.Name}");
                    }
                }

                //Assert.Inconclusive();
                ScriptControlClass.RunProcedure(gr, "ButtonShowWebsite_Click", []);
            });
        }

        public static IHostMessageBoxHostService CreateHostMessageBoxHostService()
        {
            return new TestMessageBoxHostService();
        }
        public static IHostProcessControlHostService CreateTestProcessControlHostService() // for Cli
        {
            return new TestHostProcessControlHostService();
        }

        private IHostDatabaseConnectionFactoryHostService CreateTestDatabaseConnectionFactoryHostService()
        {
            return new TestDatabaseConnectionFactoryHostService();
        }

        [TestMethod]
        public void RunTranslatedProgram()
        {
            // paste the generated .cs into gen1.cs and use the code below (adjust the 'if' and the 'external references'!)
            Type tRunner = typeof(TranslatedProgram.Runner);
            var hostServices = CreateTestHostServices();
            if (hostServices.ProvidersCount == 651) // adjust this to !=
            {
                CncIn.RunTranslatedProgram(RuntimeLogger, hostServices, TestCulture, tRunner, new Dictionary<string, object>
                 { { "Person", new DialogGuiLabelControl()} }
                , gr =>
                {
                });
            }
        }
        [TestMethod]
        public void CT132_Dialog_83() // select * from hlsysdialog where dbname = '_CustomerTest_Trumpf' and dialogid = 83; select * from hlsysdialogglobalscript where dbname = '_CustomerTest_Trumpf';
        {
            // Trumpf, 83, 83, cb_template_load_onfocus : rs.eof + rs.MoveFirst + rs.MoveNext + rs.fields("templatename").value
            // Trumpf, 78, FlagNoLicenseEndDate_ondatachange : rs.fields(0).value
            _ = nameof(DialogGuiControlBase.ControlFactoryCreateDialogControl);
            var model = new DialogGuidModel(TestCulture);
            var hlSession = new DialogGuiSession(TestCulture, agentId: 30022);

            var hlobj = new HLObjectInstance("symbol_hlobj").InitializeObjectInstance(isNew: true, TestCulture)
                    .RegisterValueKey<string>("CASEINFO.REFERENCENUMBER", 0, 0, "20260101-0001")
                    .RegisterValueKey<int>("CASEINFO.RESERVEDBY", 0, 0, 0)
                    .RegisterValueKey<string>("CaseClassificationAttribute.Impact", 0, 0, "")
                    .RegisterValueKey<string>("CaseClassificationAttribute.Priority", 0, 0, "Priority1")
                    .RegisterValueKey<string>("CaseDescription.DescriptionText", 0, 0, "")
                    .RegisterValueKey<string>("CaseDiagnosis.DiagnosisText", 0, 0, "")
                    .RegisterValueKey<string>("CaseGeneral.CostCenter", 0, 0, "cstcntr1")
                    .RegisterValueKey<string>("CaseGeneral.DefaultNotification", 0, 0, "zz1")
                    .RegisterValueKey<string>("CaseGeneral.Subject", 0, 0, "Kuku-Muku")
                    .RegisterValueKey<string>("CaseSolution.SolutionText", 0, 0, "")
                    .RegisterValueKey<string>("IncidentAttribute.Convenience", 0, 0, "")
                    .RegisterValueKey<string>("IncidentAttribute.EscalationLevel", 0, 0, "")
                    .RegisterValueKey<string>("IncidentAttribute.FunctionalRange", 0, 0, "")
                    .RegisterValueKey<string>("IncidentAttribute.IncidentStatus", 0, 0, "IncidentStatusToProof")
                    .RegisterValueKey<string>("IncidentAttribute.ProductionalRelevanz", 0, 0, "")
                    .RegisterValueKey<string>("IncidentAttribute.RequestType", 0, 0, "RequestTypeIncident")
                    .RegisterValueKey<string>("IncidentAttribute.Responsibility", 0, 0, "zz2")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailCaller", 0, 0, "")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailSearchName", 0, 0, "emsn1")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailSearchResult", 0, 0, "emsr1")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailTo", 0, 0, "emto1")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailCC", 0, 0, "emcc1")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailSubject", 0, 0, "")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailBody.TEXTVALUE", 0, 0, "")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailBody.RAWTEXT", 0, 0, "")
                    .RegisterValueKey<string>("EmailSUAttribute.EmailBody.Rawtext", 0, 0, "")
                    .RegisterValueKey<int>("Keywords.Keyword", 0, 0, 0)
                   //.RegisterValueKey<string>("CaseGeneral.DefaultNotification", 0, 0, "")
                   //.RegisterValueKey<string>("CaseGeneral.DefaultNotification", 0, 0, "")
                   //.RegisterValueKey<string>("PersonGeneral.PersonalID", 0, 0, "prsnid-x1")
                   .RegisterValueKey<string>("Keywords.KeywordOrga", 0, 0, "")
                    .RegisterValueKey<int>("SUINFO.INDEX", 0, 0, 2) // last su

                .RegisterServiceUnitIndex(1)
                    .RegisterValueKey<int>("SUINFO.INDEX", 0, 1, 1)
                    .RegisterValueKey<int>("SUINFO.EDITOR", 0, 1, 1530)


                .RegisterServiceUnitIndex(2)
                    .RegisterValueKey<int>("SUINFO.INDEX", 0, 2, 2)
                    .RegisterValueKey<int>("SUINFO.EDITOR", 0, 2, 710)
                ;
            var hlcaller = new HLObjectInstance("symbol_caller").InitializeObjectInstance(isNew: false, culture: TestCulture, objectId: 101301, objectDefName: "MyPersonDef")
                    .RegisterValueKey<string>("PersonGeneral.VIPLevel", 0, 0, "VIPLevelVIP")
                    .RegisterValueKey<string>("PersonInformation.EmailAddress", 0, 0, "zz2@svr.com")
                    .RegisterValueKey<string>("PersonGeneral.PersonSurname", 0, 0, "psn1")
                    .RegisterValueKey<string>("PersonGeneral.PersonGivenName", 0, 0, "pgn1")
                    .RegisterValueKey<string>("PersonInformation.PersonOrganisation", 0, 0, "pou1")
                    .RegisterValueKey<string>("PersonInformation.PhoneNumber", 0, 0, "ptel1")
                ;
            var hlProduct = new HLObjectInstance("hlProduct").InitializeObjectInstance(isNew: false, culture: TestCulture, objectId: null, objectDefName: "DesktopComputer")
                    .RegisterValueKey<string>("AssetGeneral.Hostname", 0, 0, "MyAN1")
                ;

            var symbol_product = new HLObjectInstance("symbol_product").InitializeObjectInstance(isNew: false, culture: TestCulture)
                    .RegisterValueKey<string>("AssetGeneral.AssetName", 0, 0, "MyAN1")
                    .RegisterValueKey<string>("AssetGeneral.Hostname", 0, 0, "MyAN1")
                    .RegisterValueKey<string>("TrumpfAssetGeneral.CINumber", 0, 0, "MyCINum1")
                ;
            model.RegisterSymbolObjectProvider("Default", () => hlobj);
            model.RegisterSymbolObjectProvider("product", () => symbol_product);
            model.RegisterSymbolObjectProvider("caller", () => hlcaller);

            IHostDatabaseConnectionFactoryHostService databaseConnectionFactoryHostService = CreateTestDatabaseConnectionFactoryHostService();

            string dialogXml = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + "_Source" + ".xml"); // CT132_Dialog_83_Source.xml
            var dialog = this.BuildDialogFromXml(dialogXml, CreateTestHostServices(services =>
            {
                services.RegisterHostService<IHostObjectFactoryHostService>(() => new TestHostObjectFactoryHostService()
                        .RegisterObjectFactory<object>("helpline.hlcontrols.HLHelperPFA", (h) => new DispatchProxyForHLHelperPFA(TestCulture))
                    );

                services.RegisterHostService<IHostMessageBoxHostService>(() => new TestMessageBoxHostService());
                services.RegisterHostService<IHostInputBoxHostService>(() => new TestInputBoxHostService());
                services.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => databaseConnectionFactoryHostService);
                services.RegisterHostService<RuntimeSupport.Implementations.FileSystemSupport.IHostFileSystemHostService>(() => new TestFileSystem()
                    .AddTestFile(@"C:\TRUMPF\helpLine\IntermediateReply.html", @"blah1")

                );

            }), model)
                .AddExternalObject("model", model, nameof(DialogGuidModel.GetPersonForAgent))
                .AddExternalObject("hlSession", hlSession)
                .AddExternalObject("hlObj", hlobj)
                .AddExternalObject("hlCaller", hlcaller)
                .AddExternalObject("hlProduct", hlProduct)
                .WorkaroundScriptCode("cb_template_load_SelectionEndOK", "position =< anzahl_agent_templates", "position <= anzahl_agent_templates") // line:1211
                .BuildDialog();

            //PERFORMANCE:for (int ixx = 1; ixx <= 7; ixx++)
            //{
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.DialogGui, dialog.ExternalReferences, isOptionalAssert: false, suppressions: ["SKY102", "SKY104", "SKY105", "SKY106", "SKY107", "SKY109"]);
            //}

            TestDialogHandlers(this, rsp, dialog, skipUnusedScript: (s) => null, s => true);
        }

        public static void TestDialogHandlers(TestBaseX tst, TestScriptResponse rsp, DialogBase dialog, Func<string, bool?> skipUnusedScript, Func<string, bool> doInvokeScript)
        {
            //string translated_cs = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension, isOptional: true);
            //if (translated_cs == null)
            //{
            //    Console.WriteLine("translating...");
            //    string scriptContent = dialog.CompleteScriptCode();
            //    translated_cs = DefaultCSharpTranslation.GetTranslatedProgramCode(TestCulture, scriptContent, dialog.ExternalReferences.Keys.ToArray());
            //}

            DoDialogGui(tst, rsp.TranslatedCsCode, dialog, (GlobalReferencesBase gr) =>
            {

                var mis = gr.GetType().GetMethods().OrderBy(x => x.Name).ToArray();
                foreach (var mi in mis)
                {
                    if (mi.DeclaringType == gr.GetType())
                    {
                        Console.WriteLine($"Method: {mi.Name}");
                    }
                }

                Dictionary<string, string> allControlEventScriptNames = new Dictionary<string, string>();
                dialog.CollectControlEventScriptNames((DialogGuiControlBase controlX, string eventNameX, string scriptNameX) =>
                {
                    if (allControlEventScriptNames.ContainsKey(scriptNameX))
                    {
                        // scriptNameX:"PriorityMatrix", "ComboBoxImpact.OnDataChange", "ComboBoxUrgency.OnDataChange"
                    }
                    else
                    {
                        allControlEventScriptNames.Add(scriptNameX, $"{controlX.ID}.{eventNameX}");
                    }
                });
                // 1: IncReqOnLoad
                string[] scriptNames = dialog.ScriptNames.OrderBy(x => x).ToArray();

                //PERFORMANCE:for (int ixx = 1; ixx <= 7; ixx++)
                {

                    int ixSearch = 0;
                    foreach (string scriptName in scriptNames)
                    {
                        if (allControlEventScriptNames.TryGetValue(scriptName, out string usedBy)
                            || string.Equals(scriptName, "OnLoad")
                            || string.Equals(scriptName, "OnSave")
                        //|| (scriptName.StartsWith("On", StringComparison.OrdinalIgnoreCase) && scriptName.Contains("_", StringComparison.OrdinalIgnoreCase))
                        )
                        {
                            // OnSave, OnLoad
                        }
                        else
                        {
                            // unused script
                            bool? skip = skipUnusedScript(scriptName);
                            if (skip == null)
                            {
                                throw new InvalidOperationException($"********** Unused script:{scriptName} *********** ");
                            }
                            else
                            {
                                if (skip.Value)
                                {
                                    Console.WriteLine($"********** Unused script:{scriptName} (SKIP)*********** ");
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($"********** Unused script:{scriptName} *********** ");
                                }
                            }
                        }

                        if (doInvokeScript != null && doInvokeScript(scriptName) == false)
                        {
                            Console.WriteLine($"NOEXEC [{scriptNames.Length}/{ixSearch + 1}] {usedBy} Invoke:{scriptName}");
                        }
                        else
                        {
                            Console.WriteLine($"[{scriptNames.Length}/{ixSearch + 1}] {usedBy} Invoke:{scriptName}");
                            //Assert.Inconclusive(); // last issue: 'IOMode' = 0 invalid argument count in 'ButtonEmailPreview_Click'
                            ScriptControlClass.RunProcedure(gr, scriptName, []);
                        }

                        ixSearch++;
                    }
                }
            });
        }

        [TestMethod]
        public void CT127_Dialog_1305()
        {
            string customerAlias = TestName.Split('_')[0];
            string dialogXml = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + "_Source" + ".xml"); // CT132_Dialog_83_Source.xml
            string customerDialogGlobalScript = TextResourceHelper.LoadResourceText<CncIn>($"Skrypton.Tests.VbsResources.{customerAlias}_DialogGlobalScript.vbs"); // see [hlsysdialogglobalscript]

            var model = new DialogGuidModel(TestCulture);
            var hlSession = new DialogGuiSession(TestCulture, agentId: 30022);

            var hlobj = new HLObjectInstance("symbol_hlobj").InitializeObjectInstance(isNew: true, culture: TestCulture)
                ;
            var hlcaller = new HLObjectInstance("symbol_caller").InitializeObjectInstance(isNew: false, culture: TestCulture, objectId: 101301, objectDefName: "MyPersonDef")
                ;

            var dialog = this.BuildDialogFromXml(dialogXml, CreateTestHostServices(services =>
                {
                    //services.RegisterHostService<IHostObjectFactoryHostService>(() => new TestHostObjectFactoryHostService()
                    //    .RegisterObjectFactory<object>("helpline.hlcontrols.HLHelperPFA", (h) => new DispatchProxyForHLHelperPFA())
                    //);

                    //services.RegisterHostService<IHostMessageBoxHostService>(() => new TestMessageBoxHostService());
                    //services.RegisterHostService<IHostInputBoxHostService>(() => new TestInputBoxHostService());
                    //services.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => databaseConnectionFactoryHostService);
                    //services.RegisterHostService<RuntimeSupport.Implementations.FileSystemSupport.IHostFileSystemHostService>(() => new TestFileSystem()
                    //    .AddTestFile(@"C:\TRUMPF\helpLine\IntermediateReply.html", @"blah1")

                    //);

                }), model)
                .AddExternalObject("model", model)
                .AddExternalObject("hlSession", hlSession)
                .AddExternalObject("hlObj", hlobj)
                .AddExternalObject("hlCaller", hlcaller)
                //.AddExternalObject("hlProduct", hlProduct)
                //.AddExternalObject("hlProduct", hlOrgunit)
                .SetGlobalScriptCode(customerDialogGlobalScript)
                .SetGlobalScriptCode(WorkaroundGlobalScript(customerAlias, customerDialogGlobalScript))
                .BuildDialog();
            model.RegisterSymbolObjectProvider("Default", () => hlobj);
            //model.RegisterSymbolObjectProvider("product", () => symbol_product);
            model.RegisterSymbolObjectProvider("caller", () => hlcaller);

            //Assert.Inconclusive(); // Compilation failed.(505,319): error CS0131: The left-hand side of an assignment must be a variable, property or indexer
            string[] suppressions = ["SKY103", "SKY104", "SKY105", "SKY106"];
            DoDialogGui(dialog, suppressions, (GlobalReferencesBase gr) =>
            {
            });
        }
        private static string WorkaroundGlobalScript(string customerAlias, string globalScript)
        {
            if (customerAlias == "CT127")
            {
                string globalScriptFixed = globalScript.Replace(
                    @"task.SetValue ""RoutingHelper.AgentID"",0,0,0,Person.GetValue(""HLOBJECTINFO.ID"",0,0,GetSvcUnitCount(),0)",
                    @"task.SetValue ""RoutingHelper.AgentID"",0,0,0,Person.GetValue(""HLOBJECTINFO.ID"",0,0,task.GetSvcUnitCount(),0)"
                );
                if (globalScriptFixed == globalScript)
                {
                    throw new NotImplementedException(customerAlias); // script not fixed
                }

                return globalScriptFixed;
            }
            return globalScript;//unchanged
        }
    }

    public static class DialogBuilderXmlExtensions
    {
        public static DialogBuilder BuildDialogFromXml(this TestBaseX tst, string dialogXml, IServiceProvider hostServices, DialogGuidModel dialogModel)
        {
            //new DialogBuilder(hostServices, "zzz").AddExternalObject("model", model)

            XElement xHelpLineDialogData = XElement.Parse(dialogXml);
            XElement xProperties = xHelpLineDialogData.Elements().Single(x => x.Name.LocalName == "Properties");
            string ObjectName = null;
            foreach (var xProperty in xProperties.Elements())
            {
                XElement xName = xProperty.Elements().Single(x => x.Name.LocalName == "Name");
                if (xName.Value == "ObjectName")
                {
                    ObjectName = xProperty.Elements().Single(x => x.Name.LocalName == "Value").Value;
                    break;
                }
            }

            Console.WriteLine($"ObjectName:{ObjectName}");
            //var xObjectName = xProperties.Elements().Single(x => x.Name.LocalName == "ObjectName").Value;

            List<DialogGuiControlBase> controls = new List<DialogGuiControlBase>();

            CollectControls(controls, xHelpLineDialogData);

            DialogBuilder builder = new DialogBuilder(hostServices, dialogModel, controls.ToArray());

            var xGuiScripts = xHelpLineDialogData.Elements().Single(x => x.Name.LocalName == "GuiScripts");

            foreach (var xScript in xGuiScripts.Elements())
            {
                string scriptName = xScript.Elements().Single(x => x.Name.LocalName == "Name").Value;
                string scriptCode = xScript.Elements().Single(x => x.Name.LocalName == "Code").Value;

                builder.AddScriptCode(scriptName, scriptCode);
            }

            return builder;
        }

        private static void CollectControls(List<DialogGuiControlBase> controls, XElement xContainer)
        {
            XElement xControls = xContainer.Elements().SingleOrDefault(x => x.Name.LocalName == "Controls");
            foreach (var xControl in xControls.Elements())
            {
                string ControlTypeName = xControl.Elements().Single(x => x.Name.LocalName == "ControlName").Value;

                //Console.WriteLine($"Read '{ControlTypeName}'");
                DialogGuiControlBase controlBase = DialogGuiControlBase.ControlFactoryCreateDialogControl(ControlTypeName); // DialogGuiGroupBox

                var xControlProperties = xControl.Elements().Single(x => x.Name.LocalName == "Properties");
                //Dictionary<string, object> controlProperties = new Dictionary<string, object>();
                foreach (var xControlProperty in xControlProperties.Elements())
                {
                    string controlPropertyName = xControlProperty.Elements().Single(x => x.Name.LocalName == "Name").Value;

                    //Console.WriteLine($"{ControlTypeName} | {controlPropertyName}");

                    var setter = controlBase.ShouldInitValueForProperty(controlPropertyName);
                    if (setter == null)
                    {
                        //Console.WriteLine($"IGNORE: {ControlTypeName} | {controlPropertyName} = ...");
                        var xName = xControlProperty.Elements().Single(x => x.Name.LocalName == "Name");
                        var xValue = xControlProperty.Elements().Single(x => x.Name.LocalName == "Value");
                        string valueTypeName = xValue.Attributes().Single(x => x.Name.LocalName == "type").Value;

                        if (valueTypeName == "HelpLineScript")
                        {
                            // <name>ButtonSubmit_Click</name>
                            string scriptName = xValue.Elements().Single(x => x.Name.LocalName == "name").Value;
                            if (!string.IsNullOrEmpty(scriptName))
                            {
                                controlBase.AddEventScript(xName.Value, scriptName);
                            }
                        }
                        else
                        {
                            //throw new NotImplementedException($"IGNORE: {ControlTypeName} | {controlPropertyName} = ({valueTypeName})");
                        }
                    }
                    else
                    {
                        // xsi:type

                        var xValue = xControlProperty.Elements().Single(x => x.Name.LocalName == "Value");
                        string valueTypeName = xValue.Attributes().Single(x => x.Name.LocalName == "type").Value;
                        object controlPropertyValue = valueTypeName switch
                        {
                            "HelpLineScript" => new HelpLineScript(xValue.Elements().Single().Value),
                            "SymbolName" => new SymbolName(xValue.Value),
                            "xsd:string" => xValue.Value,
                            "xsd:int" => XmlConvert.ToInt32(xValue.Value),
                            "xsd:boolean" => XmlConvert.ToBoolean(xValue.Value),
                            "xsd:dateTime" => XmlConvert.ToDateTime(xValue.Value, XmlDateTimeSerializationMode.Utc),
                            "ComboBoxHelplineSearch" => FromXmlComboBoxHelplineSearch(xValue),
                            _ => throw new NotImplementedException($"{ControlTypeName}.{controlPropertyName} ({valueTypeName}):{xValue.Value}")
                        };
                        if (controlPropertyName == "SymbolName" && controlPropertyValue is string sv)
                            controlPropertyValue = new SymbolName(xValue.Value);

                        //Console.WriteLine($"{ControlTypeName} | {controlPropertyName} = {controlPropertyValue}");
                        setter(controlPropertyValue);
                    }
                }

                controls.Add(controlBase);

                CollectControls(controls, xControl);
            }
        }
        private static ComboBoxHelplineSearch FromXmlComboBoxHelplineSearch(XElement xValue)
        {
            if (xValue.HasElements)
            {
                /*
                   <Value xsi:type="ComboBoxHelplineSearch">
                     <SearchObjects>
                       <string>Contact</string>
                     </SearchObjects>
                     <SearchCondition>PersonDisplayHelper.GroupIDText Like "163392*"</SearchCondition>
                     <SearchAttributeKey>PersonInformation.Name</SearchAttributeKey>
                     <AttributeKeyDefId>WorkflowHelper.AgentDefID</AttributeKeyDefId>
                     <AttributeKeyName>RoutingHelper.AgentName</AttributeKeyName>
                   </Value>
                 */
                var xSearchCondition = xValue.Elements().FirstOrDefault(x => x.Name.LocalName == "SearchCondition");
                var xAttributeKeyName = xValue.Elements().First(x => x.Name.LocalName == "AttributeKeyName");
                return new ComboBoxHelplineSearch(attributeKeyName: xAttributeKeyName.Value, searchCondition: xSearchCondition?.Value);
            }

            return new ComboBoxHelplineSearch();
        }

        internal static DialogBuilder WorkaroundScriptCode(this DialogBuilder builder, string scriptName, string originalText, string newText)
        {
            string originalCode = builder.GetScriptCode(scriptName);
            string newCode = originalCode.Replace(originalText, newText);
            if (newCode == originalCode)
                throw new InvalidOperationException("Nothing has been changed.");
            builder.FixScriptCode(scriptName, newCode);
            return builder;
        }
    }

    [DebuggerDisplay("{Name}")]
    public sealed class SymbolName
    {
        public string Name { get; }

        public SymbolName(string nameOrEmpty)
        {
            Name = nameOrEmpty;
        }

    }

    [DebuggerDisplay("{Name}")]
    internal sealed class HelpLineScript
    {
        public string Name { get; }

        public HelpLineScript(string nameOrEmpty)
        {
            Name = nameOrEmpty;
        }
    }

    internal sealed class TestHostProcessControlHostService : IHostProcessControlHostService
    {
        public TestHostProcessControlHostService()
        {

        }

        public void ProcessStart(string command, byte windowMode, bool waitOnReturn)
        {
            Console.WriteLine($"[IHostProcessControlHostService] 'ProcessStart(m:{windowMode}, w:{waitOnReturn}):' {command}");
        }

        public bool ProcessActivate(int processId)
        {
            Console.WriteLine($"[IHostProcessControlHostService] 'ProcessActivate(pid:{processId})");
            return false;
        }

        public void ProcessesCollect(Func<int, string, bool> collector)
        {
            Console.WriteLine($"[IHostProcessControlHostService] 'ProcessesCollect()");
            collector(666, "xxx");
            collector(4444, "IExplore.exe");
        }
    }

    [ComVisible(true)]
    public sealed class DialogGuidModel
    {
        private readonly CultureInfo _culture;
        public bool IsInWeb { get; private set; }

        public DialogGuidModel(CultureInfo culture, bool isInWeb = false)
        {
            _culture = culture;
            IsInWeb = isInWeb;
        }
        public void ChangeIsInWeb(bool isInWeb)
        {
            IsInWeb = isInWeb;
        }

        public void MsgBox(string message)
        {
            Console.WriteLine($"MsgBox('{message}')");
        }

        public object GetClientContext()
        {
            return null;
        }

        internal HLObjectInstance GetHelpLineTempObject(string symbolName)
        {
            if (_symbols.TryGetValue(symbolName, out var provider))
                return provider();
            return null;
        }

        internal HLObjectInstance GetHelpLineObject(string symbolName)
        {
            if (_symbols.TryGetValue(symbolName, out var provider))
                return provider();
            throw new InvalidOperationException($"symbolName:{symbolName}");
        }

        internal DialogGuidModel DialogUserControl => this;

        private Dictionary<string, Func<HLObjectInstance>> _symbols = new Dictionary<string, Func<HLObjectInstance>>(StringComparer.OrdinalIgnoreCase);
        public void RegisterSymbolObjectProvider(string symbolName, Func<HLObjectInstance> func)
        {
            _symbols.Add(symbolName, func);
        }

        private Func<HLObjectInstance> _objProvider;
        public void SetCurrentObject(Func<HLObjectInstance> objProvider)
        {
            _objProvider = objProvider ?? throw new ArgumentNullException(nameof(objProvider));
        }
        private UniqueID CurrentObjectID
        {
            get
            {
                var obj = _objProvider();
                return new UniqueID(obj.objID(), obj.TestGetObjectDefId());
            }
        }
        private sealed class UniqueID(int objID, int defID)
        {
            public int ObjID => objID;
            public int DefID => defID;
        }
        public WorkflowParameters CreateWorkflowParameters(string workFlowName)
        {
            WorkflowParameters parameters = new WorkflowParameters(workFlowName);
            parameters.Add("DEFID", this.DialogUserControl.CurrentObjectID.DefID);
            parameters.Add("OBJID", this.DialogUserControl.CurrentObjectID.ObjID);

            return parameters;
        }

        public void CreateWorkflow(string workFlowName, object parameters)
        {
            Console.WriteLine($"[DIALOGMODEL] CreateWorkflow('{workFlowName}')");
        }

        private readonly Dictionary<int, HLObjectInstance> _personOfAgents = new Dictionary<int, HLObjectInstance>();
        public DialogGuidModel InitPersonForAgent(int agentId, HLObjectInstance person)
        {
            _personOfAgents.Add(agentId, person);
            return this;
        }

        public object GetPersonForAgent(int agentId)
        {
            bool found = _personOfAgents.TryGetValue(agentId, out HLObjectInstance person);
            Console.WriteLine($"[DIALOGMODEL] GetPersonForAgent(agentId:{agentId}) : {found}");
            return found ? person : null;
        }

        public bool AllowAddNewSu { get; set; } = true;


        private readonly Dictionary<int, bool> _timerRunning = new Dictionary<int, bool>();
        public void EnableTimer(int suId, bool enable)
        {
            Console.WriteLine($"[DIALOGMODEL] EnableTimer(suId:{suId}, enable:{enable})");
            if (_timerRunning.TryGetValue(suId, out bool prevValue))
            {
                _timerRunning[suId] = enable;
            }
            else
            {
                if (enable)
                    _timerRunning.Add(suId, true);
            }
        }
        public bool IsTimerEnabled(int suId)
        {
            Console.WriteLine($"[DIALOGMODEL] IsTimerEnabled(suId:{suId})");
            return _timerRunning.TryGetValue(suId, out bool enabled) && enabled;
        }
        //[ReadOnly(true)]
        public int SaveReason
        {
            get
            {
                Console.WriteLine($"[DIALOGMODEL] SaveReason.get");
                return 0;// 0:Default,0x00000010:Associate
            }
        }
        private CurrentCommandHelper currentCommand = new CurrentCommandHelper();
        [Category("helpLine")]
        public CurrentCommandHelper CurrentCommand
        {
            get
            {
                Console.WriteLine($"[DIALOGMODEL] CurrentCommand.get");
                return currentCommand;
            }
        }

        public void SetFocusToControl(string attributeKey)
        {
            Console.WriteLine($"[DIALOGMODEL] SetFocusToControl(attributeKey:{attributeKey})");
        }

        public string Translate(string key)
        {
            Console.WriteLine($"[DIALOGMODEL] Translate(key:'{key}')");
            return $"({_culture.Name}):key";
        }

        private int _currentSUIndex;
        public int CurrentSUIndex
        {
            get
            {
                Console.WriteLine($"[DIALOGMODEL] CurrentSUIndex.get");
                return _currentSUIndex;
            }
            set
            {
                Console.WriteLine($"[DIALOGMODEL] CurrentSUIndex.set:{value}");
                if (_currentSUIndex != value)
                {
                    _currentSUIndex = value;
                }
            }
        }

        public void SetCustomerInCase(string symbol, object helpLineObject)
        {
            Console.WriteLine($"[DIALOGMODEL] SetCustomerInCase(symbol:'{symbol}, helpLineObject:{(HLObjectInstance)helpLineObject}')");
        }
    }

    [ComVisible(true)]
    public sealed class CurrentCommandHelper
    {
        public bool Aborted = false;
        public void Abort(string text)
        {
            Console.WriteLine($"[CurrentCommand] Abort(text:'{text}')");
            Aborted = true;
        }
        public void Abort()
        {
            Console.WriteLine($"[CurrentCommand] Abort()");
            Aborted = true;
        }
    }

    public class WorkflowParameters
    {
        private string workflowName;
        private Dictionary<string, object> dictionary = new Dictionary<string, object>();

        public WorkflowParameters(string workflowName)
        {
            this.workflowName = workflowName;
        }

        public void Add(string key, object value)
        {
            if (value == null || (value.GetType() != typeof(string) && value.GetType() != typeof(Int32)))
            {
                throw new Exception(String.Format("Invalid value or type given for key \"{0}\" (value: {1}, type: {2})", key, value == null ? "null" : value, value == null ? "null" : value.GetType().ToString()));
            }
            dictionary.Add(key, value);
        }
    }

    [ComVisible(true)]
    public sealed class DialogGuiSession // see public sealed class HlSession : IScriptingUserSession
    {
        private readonly CultureInfo _culture;
        private readonly int _agentId;

        public DialogGuiSession(CultureInfo culture, int agentId)
        {
            _culture = culture ?? throw new ArgumentNullException(nameof(culture));
            _agentId = agentId;
        }

        public int GetLocaleID()
        {
            Console.WriteLine($"hlSession.GetLocaleID");
            return _culture.LCID;
        }

        public int LangIDFromLCID(int lcid)
        {
            return lcid & ~1024;
        }

        public int GetAgentID()
        {
            return _agentId;
        }
    }

    static class ComInvoke
    {
        public static object Call(object comObj, string method, object[] args,
            BindingFlags flags = BindingFlags.InvokeMethod)
        {
            return comObj.GetType().InvokeMember(
                method,
                flags,
                null,
                comObj,
                args
            );
        }

        public static void SetProperty(object comObj, string prop, object value)
        {
            comObj.GetType().InvokeMember(
                prop,
                BindingFlags.SetProperty,
                null,
                comObj,
                [value]
            );
        }
    }

    // ---------------- VARIANT struct ----------------

    [StructLayout(LayoutKind.Sequential)]
    public struct VARIANT
    {
        public ushort vt;
        public ushort reserved1;
        public ushort reserved2;
        public ushort reserved3;
        public IntPtr data1;
        public IntPtr data2;

        public static VARIANT FromObject(object value)
        {
            VARIANT v = new VARIANT();

            if (value == null)
            {
                v.vt = 1; // VT_NULL
                return v;
            }

            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.String:
                    v.vt = 8; // VT_BSTR
                    v.data1 = Marshal.StringToBSTR((string)value);
                    break;

                case TypeCode.Boolean:
                    v.vt = 11; // VT_BOOL
                    v.data1 = (bool)value ? (IntPtr)(-1) : IntPtr.Zero;
                    break;

                case TypeCode.Int32:
                    v.vt = 3; // VT_I4
                    v.data1 = (IntPtr)(int)value;
                    break;

                default:
                    throw new NotSupportedException("Unsupported VARIANT type: " + value.GetType());
            }

            return v;
        }
    }
}