using System;
using System.Collections.Generic;
using System.Text;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// An offline IHostHttpClientHostService for the tests: every request is answered with status 200 and SampleResponseXml (no network access).
    /// The requests are recorded.
    /// </summary>
    internal sealed class TestHttpClientHostService : IHostHttpClientHostService
    {
        public List<HostHttpRequest> Requests { get; } = new List<HostHttpRequest>();

        public HostHttpResponse Send(HostHttpRequest request)
        {
            Requests.Add(request ?? throw new ArgumentNullException(nameof(request)));
            return new HostHttpResponse(
                200,
                "OK",
                new[] { new KeyValuePair<string, string>("Content-Type", "text/xml; charset=utf-8") },
                Encoding.UTF8.GetBytes(SampleResponseXml)
            );
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
    }
}
