Public Function ConvertSize(ByRef Size)

  'MsgBox "Converting Size for " & Size
  Size = CSng(Replace(Size, ",", ""))

  IF Not VarType(Size) = vbSingle THEN
    ConvertSize = "SIZE INPUT ERROR"
    Exit Function
  END IF

  Suffix = " B"
  IF Size >= 1024 THEN
    suffix = " KB"
  END IF
  IF Size >= 1048576 THEN
    suffix = " MB"
  END IF
  IF Size >= 1073741824 THEN
    suffix = " GB"
  END IF
  IF Size >= 1099511627776 THEN
    suffix = " TB"
  END IF

  SELECT CASE Suffix
    CASE " KB"
      Size = Round(Size / 1024, 2)
    CASE " MB"
      Size = Round(Size / 1048576, 2)
    CASE " GB"
      Size = Round(Size / 1073741824, 2)
    CASE " TB"
      Size = Round(Size / 1099511627776, 2)
  END SELECT

  ConvertSize = Size & Suffix
End Function
Public Function getNexthinkUser()
  getNexthinkUser = "myusr2"
End Function
Public Function getNexthinkBaseURL()
  getNexthinkBaseURL = ""
End Function
Public Function getNexthinkPassword()
  getNexthinkPassword = "mypwd2"
End Function
Public Sub ButtonGeneralInfo_Click()

  TabPageGeneralInfo.ShowControl = 1
  TabPageSoftwareOSHealth.ShowControl = 3
  TabPageSecurityCompliance.ShowControl = 3
  TabPageTechnicalInfo.ShowControl = 3
  TabPageNetworkHealth.ShowControl = 3
  TabPageL1Checklist.ShowControl = 3

  GroupBoxGeneralInfo.ShowControl = 1
  GroupBoxTechnicalInfo.ShowControl = 3
  GroupBoxSoftwareOSHealth.ShowControl = 3
  GroupBoxSecurityCompliance.ShowControl = 3
  GroupBoxNetworkHealth.ShowControl = 3
  GroupBoxL1Checklist.ShowControl = 3

  ButtonGeneralInfo.BackColor = "#5b5b5b"
  ButtonTechnicalInfo.BackColor = "#1B709F"
  ButtonSWHealth.BackColor = "#1B709F"
  ButtonSecurityCompliance.BackColor = "#1B709F"
  ButtonNetworkHealth.BackColor = "#1B709F"
  ButtonL1Checklist.BackColor = "#1B709F"

  TabControlNexthink.ShowControl = 1
  TabPageGeneralInfo.RequestFocus = True

  'Clear TextBoxes

  TextBoxGeneralCallTime.Text = ""
  TextBoxGeneralHostName.Text = ""
  TextBoxlGeneralDeviceManufacturer.Text = ""
  TextBoxGeneralDeviceProductVersion.Text = ""
  TextBoxGeneralLastIP.Text = ""
  TextBoxGeneralGroupName.Text = ""
  TextBoxGeneralOS.Text = ""
  TextBoxGeneralLastBootTime.Text = ""
  TextBoxGeneralLastLogon.Text = ""
  TextBoxGeneralDeviceType.Text = ""
  TextBoxGeneralBIOSSerialNumber.Text = ""
  TextBoxGeneralCPUModel.Text = ""
  TextBoxGeneralNumberOfCPUs.Text = ""
  TextBoxGeneralNumberOfLogProcs.Text = ""
  TextBoxGeneralNumberOfCores.Text = ""
  TextBoxGeneralCPUFreq.Text = ""
  TextBoxGeneralTotalRAM.Text = ""
  TextBoxGeneralNumberOfGraphCards.Text = ""

  ' --- GroupBoxTechnicalInfo

  TextBoxTechnicalInfoTotalDriveCapNow.Text = ""
  TextBoxTechnicalInfoTotalFreeSpaceNow.Text = ""
  TextBoxTechnicalInfoTotalDriveUsageNow.Text = ""
  TextBoxTechnicalInfoSystemDriveCapNow.Text = ""
  TextBoxTechnicalInfoSystemDriveFreeSpaceNow.Text = ""
  TextBoxTechnicalInfoHighCPUTimeNow.Text = ""
  TextBoxTechnicalInfoHighMemoryTimeNow.Text = ""
  TextBoxTechnicalInfoHighIOTimeNow.Text = ""
  TextBoxTechnicalInfoTotalDriveCap7Days.Text = ""
  TextBoxTechnicalInfoTotalFreeSpace7Days.Text = ""
  TextBoxTechnicalInfoTotalDriveUsage7Days.Text = ""
  TextBoxTechnicalInfoSystemDriveCap7Days.Text = ""
  TextBoxTechnicalInfoSystemDriveFreeSpace7Days.Text = ""
  TextBoxTechnicalInfoHighCPUTime7Days.Text = ""
  TextBoxTechnicalInfoHighMemoryTime7Days.Text = ""
  TextBoxTechnicalInfoHighIOTime7Days.Text = ""

  ImageNOKTechnicalInfoTotalFreeSpaceNow.ShowControl = 3
  ImageOKTechnicalInfoTotalFreeSpaceNow.ShowControl = 3
  ImageNOKTechnicalInfoTotalDriveUsageNow.ShowControl = 3
  ImageOKTechnicalInfoTotalDriveUsageNow.ShowControl = 3
  ImageNOKTechnicalInfoSystemDriveCapNow.ShowControl = 3
  ImageOKTechnicalInfoSystemDriveCapNow.ShowControl = 3
  ImageNOKTechnicalInfoSystemDriveFreeSpaceNow.ShowControl = 3
  ImageOKTechnicalInfoSystemDriveFreeSpaceNow.ShowControl = 3
  ImageNOKTechnicalInfoHighCPUTimeNow.ShowControl = 3
  ImageOKTechnicalInfoHighCPUTimeNow.ShowControl = 3
  ImageNOKTechnicalInfoHighMemoryTimeNow.ShowControl = 3
  ImageOKTechnicalInfoHighMemoryTimeNow.ShowControl = 3
  ImageNOKTechnicalInfoHighIOTimeNow.ShowControl = 3
  ImageOKTechnicalInfoHighIOTimeNow.ShowControl = 3
  ImageOKTechnicalInfoTotalFreeSpace7Days.ShowControl = 3
  ImageOKTechnicalInfoTotalFreeSpace7Days.ShowControl = 3
  ImageNOKTechnicalInfoTotalDriveUsage7Days.ShowControl = 3
  ImageOKTechnicalInfoTotalDriveUsage7Days.ShowControl = 3
  ImageNOKTechnicalInfoSystemDriveCap7Days.ShowControl = 3
  ImageOKTechnicalInfoSystemDriveCap7Days.ShowControl = 3
  ImageNOKTechnicalInfoSystemDriveFreeSpace7Days.ShowControl = 3
  ImageOKTechnicalInfoSystemDriveFreeSpace7Days.ShowControl = 3
  ImageNOKTechnicalInfoHighCPUTime7Days.ShowControl = 3
  ImageOKTechnicalInfoHighCPUTime7Days.ShowControl = 3
  ImageNOKTechnicalInfoHighMemoryTime7Days.ShowControl = 3
  ImageOKTechnicalInfoHighMemoryTime7Days.ShowControl = 3
  ImageNOKTechnicalInfoHighIOTime7Days.ShowControl = 3
  ImageOKTechnicalInfoHighIOTime7Days.ShowControl = 3

  ' GroupBox Software OS Health

  TextBoxSoftwareOSHealthOSVersionArchitecture.Text = ""
  TextBoxSoftwareOSHealthOSName.Text = ""
  TextBoxSoftwareOSHealthWMIStatus.Text = ""
  TextBoxSoftwareOSHealthLastSystemUpdate.Text = ""
  TextBoxSoftwareOSHealthWindowsUpdateStatus.Text = ""
  TextBoxSoftwareOSHealthNumberOfApps.Text = ""
  TextBoxSoftwareOSHealthNumberOfExes.Text = ""
  TextBoxSoftwareOSHealthNumberOfBins.Text = ""
  TextBoxSoftwareOSHealthOSEndOfSupport.Text = ""
  TextBoxSoftwareOSHealthOSIE11Support.Text = ""
  TextBoxSoftwareOSHealthWin10Ready.Text = ""
  TextBoxSoftwareOSHealthOSComplience.Text = ""

  ' GroupBox Security Compliance

  TextBoxSecurityComplianceInetSecuritySettings.Text = ""
  TextBoxSecurityComplianceUserAccountStatus.Text = ""
  TextBoxSecurityComplianceAntivirusName.Text = ""
  TextBoxSecurityComplianceAntivirusRTP.Text = ""
  TextBoxSecurityComplianceAntivirusUpToDate.Text = ""
  TextBoxSecurityComplianceAntivirusNumber.Text = ""
  TextBoxSecurityComplianceAntivirusAll.Text = ""
  TextBoxSecurityComplianceAntispywareName.Text = ""
  TextBoxSecurityComplianceAntispywareRTP.Text = ""
  TextBoxSecurityComplianceAntispywareUpToDate.Text = ""
  TextBoxSecurityComplianceAntispywareNumber.Text = ""
  TextBoxSecurityComplianceAntispywareAll.Text = ""
  TextBoxSecurityComplianceFirewallName.Text = ""
  TextBoxSecurityComplianceFirewallRTP.Text = ""
  TextBoxSecurityComplianceFirewallNumber.Text = ""
  TextBoxSecurityComplianceFirewallAll.Text = ""

  ' GroupBox Network Health
  TextBoxNetworkHealthIncomingNetTaffic24Hours.Text = ""
  TextBoxNetworkHealthOutgoingNetTaffic24Hours.Text = ""
  TextBoxNetworkHealthTotalNetTaffic24Hours.Text = ""
  TextBoxNetworkHealthSuccessNetConnectionRatio24Hours.Text = ""
  TextBoxNetworkHealthNetAvailLevel24Hours.Text = ""
  TextBoxNetworkHealthAvgIncomingNetBitrate24Hours.Text = ""
  TextBoxNetworkHealthAvgOutgoingNetBitrate24Hours.Text = ""
  TextBoxNetworkHealthAvgNetResponseTime24Hours.Text = ""
  TextBoxNetworkHealthIncomingWebTraffic24Hours.Text = ""
  TextBoxNetworkHealthOutgoingWebTraffic24Hours.Text = ""
  TextBoxNetworkHealthTotalWebTraffic24Hours.Text = ""
  TextBoxNetworkHealthAvgIncomingWebBitrate24Hours.Text = ""
  TextBoxNetworkHealthAvgOutgoingWebBitrate24Hours.Text = ""
  TextBoxNetworkHealthAvgWebRequestSize24Hours.Text = ""
  TextBoxNetworkHealthAvgWebResponseSize24Hours.Text = ""
  TextBoxNetworkHealthSuccessHTTPRequestRatio24Hours.Text = ""

  TextBoxNetworkHealthIncomingNetTaffic7Days.Text = ""
  TextBoxNetworkHealthOutgoingNetTaffic7Days.Text = ""
  TextBoxNetworkHealthTotalNetTaffic7Days.Text = ""
  TextBoxNetworkHealthSuccessNetConnectionRatio7Days.Text = ""
  TextBoxNetworkHealthNetAvailLevel7Days.Text = ""
  TextBoxNetworkHealthAvgIncomingNetBitrate7Days.Text = ""
  TextBoxNetworkHealthAvgOutgoingNetBitrate7Days.Text = ""
  TextBoxNetworkHealthAvgNetResponseTime7Days.Text = ""
  TextBoxNetworkHealthIncomingWebTraffic7Days.Text = ""
  TextBoxNetworkHealthOutgoingWebTraffic7Days.Text = ""
  TextBoxNetworkHealthTotalWebTraffic7Days.Text = ""
  TextBoxNetworkHealthAvgIncomingWebBitrate7Days.Text = ""
  TextBoxNetworkHealthAvgOutgoingWebBitrate7Days.Text = ""
  TextBoxNetworkHealthAvgWebRequestSize7Days.Text = ""
  TextBoxNetworkHealthAvgWebResponseSize7Days.Text = ""
  TextBoxNetworkHealthSuccessHTTPRequestRatio7Days.Text = ""

  ImageOKNetworkHealthIncomingNetTaffic24Hours.ShowControl = 3
  ImageNOKNetworkHealthIncomingNetTaffic24Hours.ShowControl = 3
  ImageOKNetworkHealthIncomingNetTaffic7Days.ShowControl = 3
  ImageNOKNetworkHealthIncomingNetTaffic7Days.ShowControl = 3
  ImageOKNetworkHealthOutgoingNetTaffic24Hours.ShowControl = 3
  ImageNOKNetworkHealthOutgoingNetTaffic24Hours.ShowControl = 3
  ImageOKNetworkHealthOutgoingNetTaffic7Days.ShowControl = 3
  ImageNOKNetworkHealthOutgoingNetTaffic7Days.ShowControl = 3
  ImageOKNetworkHealthTotalNetTaffic24Hours.ShowControl = 3
  ImageNOKNetworkHealthTotalNetTaffic24Hours.ShowControl = 3
  ImageOKNetworkHealthTotalNetTaffic7Days.ShowControl = 3
  ImageNOKNetworkHealthTotalNetTaffic7Days.ShowControl = 3
  ImageOKNetworkHealthSuccessNetConnectionRatio24Hours.ShowControl = 3
  ImageNOKNetworkHealthSuccessNetConnectionRatio24Hours.ShowControl = 3
  ImageOKNetworkHealthSuccessNetConnectionRatio7Days.ShowControl = 3
  ImageNOKNetworkHealthSuccessNetConnectionRatio7Days.ShowControl = 3
  ImageOKNetworkHealthNetAvailLevel24Hours.ShowControl = 3
  ImageNOKNetworkHealthNetAvailLevel24Hours.ShowControl = 3
  ImageOKNetworkHealthNetAvailLevel7Days.ShowControl = 3
  ImageNOKNetworkHealthNetAvailLevel7Days.ShowControl = 3
  ImageOKNetworkHealthAvgIncomingNetBitrate24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgIncomingNetBitrate24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgIncomingNetBitrate7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgIncomingNetBitrate7Days.ShowControl = 3
  ImageOKNetworkHealthAvgOutgoingNetBitrate24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgOutgoingNetBitrate24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgOutgoingNetBitrate7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgOutgoingNetBitrate7Days.ShowControl = 3
  ImageOKNetworkHealthAvgNetResponseTime24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgNetResponseTime24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgNetResponseTime7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgNetResponseTime7Days.ShowControl = 3
  ImageOKNetworkHealthIncomingWebTraffic24Hours.ShowControl = 3
  ImageNOKNetworkHealthIncomingWebTraffic24Hours.ShowControl = 3
  ImageOKNetworkHealthIncomingWebTraffic7Days.ShowControl = 3
  ImageNOKNetworkHealthIncomingWebTraffic7Days.ShowControl = 3
  ImageOKNetworkHealthOutgoingWebTraffic24Hours.ShowControl = 3
  ImageNOKNetworkHealthOutgoingWebTraffic24Hours.ShowControl = 3
  ImageOKNetworkHealthOutgoingWebTraffic7Days.ShowControl = 3
  ImageNOKNetworkHealthOutgoingWebTraffic7Days.ShowControl = 3
  ImageOKNetworkHealthTotalWebTraffic24Hours.ShowControl = 3
  ImageOKNetworkHealthTotalWebTraffic24Hours.ShowControl = 3
  ImageOKNetworkHealthTotalWebTraffic7Days.ShowControl = 3
  ImageNOKNetworkHealthTotalWebTraffic7Days.ShowControl = 3
  ImageOKNetworkHealthAvgIncomingWebBitrate24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgIncomingWebBitrate24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgIncomingWebBitrate7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgIncomingWebBitrate7Days.ShowControl = 3
  ImageOKNetworkHealthAvgOutgoingWebBitrate24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgOutgoingWebBitrate24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgOutgoingWebBitrate7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgOutgoingWebBitrate7Days.ShowControl = 3
  ImageOKNetworkHealthAvgWebRequestSize24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgWebRequestSize24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgWebRequestSize7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgWebRequestSize7Days.ShowControl = 3
  ImageOKNetworkHealthAvgWebResponseSize24Hours.ShowControl = 3
  ImageNOKNetworkHealthAvgWebResponseSize24Hours.ShowControl = 3
  ImageOKNetworkHealthAvgWebResponseSize7Days.ShowControl = 3
  ImageNOKNetworkHealthAvgWebResponseSize7Days.ShowControl = 3
  ImageOKNetworkHealthSuccessHTTPRequestRatio24Hours.ShowControl = 3
  ImageNOKNetworkHealthSuccessHTTPRequestRatio24Hours.ShowControl = 3
  ImageOKNetworkHealthSuccessHTTPRequestRatio7Days.ShowControl = 3
  ImageNOKNetworkHealthSuccessHTTPRequestRatio7Days.ShowControl = 3

  ' GroupBox L1-Checkliste

  TextBoxL1FreeSpace.Text = ""
  TextBoxL1OSUpToDate.Text = ""
  TextBoxL1Browser.Text = ""
  TextBoxL1Collaboration.Text = ""
  TextBoxL1Antivirus.Text = ""
  TextBoxL1Antivirus2.Text = ""
  TextBoxL1Antivirus3.Text = ""
  TextBoxL1Defender.Text = ""
  TextBoxL1BootLogon2.Text = ""
  TextBoxL1BootLogon3.Text = ""
  TextBoxL1CPU24.Text = ""
  TextBoxL1CPU7.Text = ""
  TextBoxL1Speicher24.Text = ""
  TextBoxL1Speicher7.Text = ""
  TextBoxL1Bluescreen24.Text = ""
  TextBoxL1Bluescrren7.Text = ""
  TextBoxL1HardReset24.Text = ""
  TextBoxL1HardReset7.Text = ""


  ' --- GroupBoxGeneralInfo

  Dim nexthinkBaseURL
  nexthinkBaseURL = getNexthinkBaseURL() & "query?p1="
  Dim nexthinkQuery
  nexthinkQuery = "&platform=windows&query=(select (name last_ip_address group_name last_logged_on_user os_version_and_architecture device_manufacturer number_of_cpus cpu_model number_of_cores logical_cpu_number cpu_frequency total_ram number_of_graphical_cards graphical_card_ram last_system_boot last_logon_time bios_serial_number device_model ) (from device (where device (eq name (string %1))) ))&format=xml"
  Dim nexthinkURL

  Dim colorWarning
  colorWarning = "#F20012"
  Dim colorCheck
  colorCheck = "#1B709F"

  Dim hostname
  hostname = hlObj.GetValue("ComputerDetail.Hostname", 0, 0, 0, 0)

  IF hostname = "" THEN
    model.MsgBox "Der Computer hat keinen Hostnamen."
    Exit Sub
  END IF

  nexthinkURL = nexthinkBaseURL & UCase(hostname) & nexthinkQuery
  nexthinkURL = "https://httpbin.org/get"

  ON ERROR RESUME NEXT


  'MsgBox nexthinkURL

  'time of call
  TextBoxGeneralCallTime.Text = FormatDateTime(Now, vbGeneralDate)

  Dim xmlhttp
  Set xmlhttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
  xmlhttp.setOption 2, 13056
  'bypass certificate errors
  xmlhttp.open "GET", nexthinkURL, False, getNexthinkUser(), getNexthinkPassword()
  xmlhttp.send

  'Error Handling
  IF Err.Number <> 0 THEN
    model.MsgBox "Beim Nexthink Abruf (POST) ist ein Fehler aufgetreten. Möglicherweise ist der Server nicht erreichbar."
    model.MsgBox "Error Description: " & Err.Description & vbLf & "Error Source: " & Err.Source & vbLf & "Error HelpFile: " & Err.Helpfile & vbLf & "Error Context: " & Err.HelpContext
    Exit Sub
  END IF

  'Reset the Error Data
  Err.Clear

  Set xmlDoc = CreateObject("Msxml2.DOMDocument")
  xmlDoc.async = "false"
  xmlDoc.load(xmlhttp.responseXML)

  'Error Handling
  IF Err.Number <> 0 THEN
    model.MsgBox "Beim Nexthink Abruf (GET) ist ein Fehler aufgetreten."
    model.MsgBox "Error Description: " & Err.Description & vbLf & "Error Source: " & Err.Source & vbLf & "Error HelpFile: " & Err.Helpfile & vbLf & "Error Context: " & Err.HelpContext
    Exit Sub
  END IF


  Dim dict
  Set dict = CreateObject("Scripting.Dictionary")

  Dim curnode

  'iterate all nodes and write into dictionary
  Dim i
  i = 0
  For Each n In xmlDoc.SelectNodes("//table/header/*")
    Set curnode = xmlDoc.documentElement.selectSingleNode("//table/body/r/c" & i)
    dict.Add n.Text, curnode.Text
    i = i + 1
  Next

  'Error Handling
  IF Err.Number <> 0 THEN
    model.MsgBox "Beim Verarbeiten der Nexthink Informationen ist ein Fehler aufgetreten."
    Exit Sub
  END IF

  ' from now on ->; possibility to access dictionary by dict.Item("KEY") KEY = name of node

  'fill textboxes
  'LabelNName.Text = dict.key("name")
  TextBoxGeneralHostName.Text = dict.Item("name")
  TextBoxGeneralLastIP.Text = dict.Item("last_ip_address")
  TextBoxlGeneralDeviceManufacturer.Text = dict.Item("device_manufacturer")
  TextBoxGeneralDeviceProductVersion.Text = dict.Item("device_model")
  TextBoxGeneralOS.Text = dict.Item("os_version_and_architecture")
  TextBoxGeneralGroupName.Text = dict.Item("group_name")
  TextBoxGeneralLastBootTime.Text = FormatDateTime(Replace(dict.Item("last_system_boot"), "T", " "), vbGeneralDate)
  TextBoxGeneralLastLogon.Text = FormatDateTime(Replace(dict.Item("last_logon_time"), "T", " "), vbGeneralDate)
  TextBoxGeneralDeviceType.Text = dict.Item("last_logged_on_user")
  TextBoxGeneralBIOSSerialNumber.Text = dict.Item("bios_serial_number")
  TextBoxGeneralCPUModel.Text = dict.Item("cpu_model")
  TextBoxGeneralNumberOfCPUs.Text = dict.Item("number_of_cpus")
  TextBoxGeneralNumberOfLogProcs.Text = dict.Item("logical_cpu_number")
  TextBoxGeneralNumberOfCores.Text = dict.Item("number_of_cores")
  TextBoxGeneralCPUFreq.Text = dict.Item("cpu_frequency") & " MHz"
  TextBoxGeneralTotalRAM.Text = ConvertSize(dict.Item("total_ram"))

  TextBoxGeneralNumberOfGraphCards.Text = dict.Item("number_of_graphical_cards")
  TextBoxGeneralGraphCardRAM.Text = ConvertSize(dict.Item("graphical_card_ram"))

End Sub
ButtonGeneralInfo_Click
