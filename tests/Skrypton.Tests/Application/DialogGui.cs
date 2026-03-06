using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class DialogGui : TestBase
    {
        [TestMethod]
        public void QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click()// => TestDialogGui();
        //private void TestDialogGui()
        {
            var dialog = new DialogBuilder(CreateTestHostServices())
                .AddTextControl("TextBoxWebsite")
                .BuildDialog();
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(dialog, gr => { });
        }

        [TestMethod]
        public void CT35_LogChecklist_Dialog_388_OnSave() // 35:DFSnDLNeu  id = select id, dbname from _databasestats order by dbname asc -- [hlsysdialog]
        {
            var dialog = new DialogBuilder(CreateTestHostServices())
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

            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(dialog, gr => { });

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
            Test_IDispatch_Invoke(); //works only on windows

            var hlobj = new HLObjectInstance()
                    .RegisterValueKey<string>("ComputerDetail.Hostname", 0, 0, "hst-X_1")
                ;
            var model = new DialogGuidModel();

            var dialog = new DialogBuilder(CreateTestHostServices()).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
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


            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(dialog, gr => { });
        }

        private void DoDialogGui(DialogBase dialog, Action<GlobalReferencesBase> dialogHandler)
        {
            CncIn.ExecuteTranslatedProgram(RuntimeLogger, dialog.HostServices, TestCulture, TestContext.TestName, dialog.ExternalReferences, dialogHandler);
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
            var model = new DialogGuidModel();

            IHostProcessControlHostService processControlHostService = CreateTestProcessControlHostService();

            var dialog = new DialogBuilder(CreateTestHostServices(r => r.RegisterHostService<IHostProcessControlHostService>(() => processControlHostService))).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
                .AddButton("Button1_Click")
                .BuildDialog();

            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(dialog, (GlobalReferencesBase gr) =>
            {
                var mis = gr.GetType().GetMethods().OrderBy(x => x.Name).ToArray();
                foreach (var mi in mis)
                {
                    if (mi.DeclaringType == gr.GetType())
                    {
                        Console.WriteLine($"Method: {mi.Name}");
                    }
                }

                ScriptControlClass.RunProcedure(gr, "Button1_click", []);
            });
        }

        [TestMethod]
        public void CT74_ClientComputer_Dialog_2_ButtonShowWebsite_Click() // select * from hlsysdialog where dbname = '_CustomerTest_Mainova' and dialogid = 2; select * from hlsysdialogglobalscript where dbname = '_CustomerTest_Mainova';
        {
            var hlobj = new HLObjectInstance().InitializeObjectInstance(isNew: true)
                    .RegisterValueKey<string>("PersonBilling.CostCenter_CA", 0, 0, "hst-X_1")
                    .RegisterValueKey<string>("PersonInformation.SBCode", 0, 0, "hst-X_1")
                    .RegisterValueKey<string>("PersonGeneral.Name", 0, 0, "Kuku-Muku")
                    .RegisterValueKey<string>("PersonGeneral.Group", 0, 0, "g-x1")
                    .RegisterValueKey<string>("PersonGeneral.PersonalID", 0, 0, "prsnid-x1")
                ;
            var model = new DialogGuidModel();

            IHostDatabaseConnectionFactoryHostService databaseConnectionFactoryHostService = CreateTestDatabaseConnectionFactoryHostService();

            var dialog = new DialogBuilder(CreateTestHostServices(r => r.RegisterHostService<IHostDatabaseConnectionFactoryHostService>(() => databaseConnectionFactoryHostService))).AddExternalObject("model", model).AddExternalObject("hlobj", hlobj)
                .AddButton("ButtonShowWebsite_Click")
                .BuildDialog();

            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, dialog.ExternalReferences);
            DoDialogGui(dialog, (GlobalReferencesBase gr) =>
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

        private IHostProcessControlHostService CreateTestProcessControlHostService()
        {
            return new TestHostProcessControlHostService();
        }

        private IHostDatabaseConnectionFactoryHostService CreateTestDatabaseConnectionFactoryHostService()
        {
            return new TestDatabaseConnectionFactoryHostService();
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
