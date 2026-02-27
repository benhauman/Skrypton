using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    // "Msxml2.ServerXMLHTTP.6.0"
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    //[Guid("F5078F35-C551-11D3-89B9-0000F81FE221")]   // CLSID_ServerXMLHTTP60
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal sealed class MyServerXMLHTTP60 : IDispatchBase.IDispatchOnReflect//IDispatchBase // : IServerXMLHTTP60
    {
        /*
            Summary Table
            -----------------------------------------
            ProgID (versioned):       Msxml2.ServerXMLHTTP.6.0
            ProgID (unversioned):     Msxml2.ServerXMLHTTP
            CLSID:                    {F5078F35-C551-11D3-89B9-0000F81FE221}
            COM Class Name:           ServerXMLHTTP60
            DLL:                      msxml6.dll

            Additional options available:
              - List all MSXML ProgIDs and class names
              - Show the exact registry layout for each MSXML COM class
              - Demonstrate how to instantiate via CLSID instead of ProgID
              - Show how to enumerate all MSXML COM classes programmatically
        */

        public MyServerXMLHTTP60() : base(typeof(MyServerXMLHTTP60))
        {
        }

        private sealed class MyRequest
        {
            public string method { get; set; }
            public string url { get; set; }
            public bool isasync { get; set; }
            public string user { get; set; }
            public string password { get; set; }
            public object body { get; set; }
        }
        private MyRequest _request;

        // void open(BSTR method, BSTR url, VARIANT async, VARIANT user, VARIANT password)
        [DispId(0x00000001)]
        public void open(
            string method,
            string url,
            object async = null,
            object user = null,
            object password = null)
        {
            _request = new MyRequest()
            {
                method = method,
                url = url,
                isasync = (bool)ToBool(async),
                user = (string)user,
                password = (string)password
            };
        }

        private static bool ToBool(object value)
        {
            if (value is short shortValue)
                return shortValue != 0;
            if (value is int intValue)
                return intValue != 0;
            return (bool)value;
        }

        // void setRequestHeader(BSTR header, BSTR value)
        [DispId(0x00000002)]
        public void setRequestHeader(string header, string value)
        {
            throw new NotImplementedException();
        }

        // BSTR getResponseHeader(BSTR header)
        [DispId(0x00000003)]
        public string getResponseHeader(string header)
        {
            throw new NotImplementedException();
        }

        // BSTR getAllResponseHeaders()
        [DispId(0x00000004)]
        public string getAllResponseHeaders()
        {
            throw new NotImplementedException();
        }

        public const string SampleResponseXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<table>
  <header>
    <name>name</name>
    <last_ip_address>last_ip_address</last_ip_address>
    <group_name>group_name</group_name>
    <last_logged_on_user>last_logged_on_user</last_logged_on_user>
    <os_version_and_architecture>os_version_and_architecture</os_version_and_architecture>
    <device_manufacturer>device_manufacturer</device_manufacturer>
    <number_of_cpus>number_of_cpus</number_of_cpus>
    <cpu_model>cpu_model</cpu_model>
    <number_of_cores>number_of_cores</number_of_cores>
    <logical_cpu_number>logical_cpu_number</logical_cpu_number>
    <cpu_frequency>cpu_frequency</cpu_frequency>
    <total_ram>total_ram</total_ram>
    <number_of_graphical_cards>number_of_graphical_cards</number_of_graphical_cards>
    <graphical_card_ram>graphical_card_ram</graphical_card_ram>
    <last_system_boot>last_system_boot</last_system_boot>
    <last_logon_time>last_logon_time</last_logon_time>
    <bios_serial_number>bios_serial_number</bios_serial_number>
    <device_model>device_model</device_model>
  </header>

  <body>
    <r>
      <c0>LAPTOP-123</c0>
      <c1>192.168.10.45</c1>
      <c2>Finance</c2>
      <c3>DOMAIN\jdoe</c3>
      <c4>Windows 11 Pro x64</c4>
      <c5>Dell Inc.</c5>
      <c6>1</c6>
      <c7>Intel(R) Core(TM) i7-1185G7</c7>
      <c8>4</c8>
      <c9>8</c9>
      <c10>3000</c10>
      <c11>16384</c11>
      <c12>1</c12>
      <c13>4096</c13>
      <c14>2025-01-04T08:12:30Z</c14>
      <c15>2025-01-05T09:01:15Z</c15>
      <c16>ABC123XYZ</c16>
      <c17>Latitude 7420</c17>
    </r>
  </body>
</table>

";

        [DispId(0x00000005)]
        public void send()
        {
            send(null);
        }
        // void send(VARIANT body)
        [DispId(0x00000005)]
        public void send(object body) // body can be null
        {
            _request.body = body;
            _response = new MyResponse() { responseXML = SampleResponseXml };
        }
        private MyResponse _response;
        class MyResponse
        {
            public string responseXML { get; set; }
            public string responseText { get; set; }
            public int status { get; set; }
        }

        // IDispatch responseXML
        [DispId(0x00000006)]
        public object responseXML
        {
            get
            {
                byte[] xmlBytes = Encoding.UTF8.GetBytes(_response.responseXML);
                return xmlBytes;
            }
        }

        // BSTR responseText
        [DispId(0x00000007)]
        public string responseText
        {
            get
            {
                return _response.responseText;
            }
        }

        // long status
        // it is a standard HTTP status code (e.g., 200, 404, 500, etc.).
        // Represents the HTTP status code returned by a request.
        // The results of this method are valid only after the send method has been successfully completed.
        [DispId(0x00000008)]
        public int status
        {
            get
            {
                return _response.status;
            }
        }

        // BSTR statusText
        [DispId(0x00000009)]
        public string statusText { get { throw new NotImplementedException(); } }

        // void abort()
        [DispId(0x0000000A)]
        public void abort()
        {
            throw new NotImplementedException();
        }

        // void setTimeouts(long resolve, long connect, long send, long receive)
        [DispId(0x0000000B)]
        public void setTimeouts(int resolve, int connect, int send, int receive)
        {
            throw new NotImplementedException();
        }

        private readonly Dictionary<int, object> _options = new Dictionary<int, object> { { 2, "2" } };
        // void setOption(int option, VARIANT value)
        [DispId(0x0000000C)]
        public void setOption(int option, object value)
        {
            _options[option] = value;
        }
        // VARIANT getOption(int option)
        [DispId(0x0000000D)]
        public object getOption(int option)
        {
            return _options[option];
        }
    }
}
