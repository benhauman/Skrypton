' Demo: compose an Outlook mail (recipients, HTML body, attachment), a contact and an appointment.
'   Skrypton.Windows.Cli Outlook-Mail.vbs           shows the mail in Outlook (Display) - nothing is sent
'   Skrypton.Windows.Cli Outlook-Mail.vbs /dryrun   only prints what it composed
Option Explicit
Dim dryRun, fso, file, stream, ol, mail, contact, appointment
dryRun = False
If WScript.Arguments.Count > 0 Then dryRun = (LCase(WScript.Arguments(0)) = "/dryrun")

Set fso = CreateObject("Scripting.FileSystemObject")
file = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-demo-report.txt")
Set stream = fso.CreateTextFile(file, True)
stream.WriteLine "Monthly report"
stream.Close

Set ol = CreateObject("Outlook.Application")
Set mail = ol.CreateItem(0) ' olMailItem
mail.Subject = "Monthly report"
mail.To = "team@example.invalid; lead@example.invalid"
mail.CC = "boss@example.invalid"
mail.HTMLBody = "<p>Hello team,</p><p>please find the report attached.</p>"
mail.Attachments.Add file
WScript.Echo "Mail: " & mail.Subject
WScript.Echo "  To: " & mail.To & " | CC: " & mail.CC & " | recipients: " & mail.Recipients.Count
WScript.Echo "  attachments: " & mail.Attachments.Count & " | HTML: " & (mail.BodyFormat = 2)

Set contact = ol.CreateItem(2) ' olContactItem
contact.FirstName = "Max"
contact.LastName = "Muster"
contact.CompanyName = "Example Ltd"
contact.Email1Address = "max@example.invalid"
WScript.Echo "Contact: " & contact.FullName & " (" & contact & "), " & contact.CompanyName & ", " & contact.Email1Address

Set appointment = ol.CreateItem(1) ' olAppointmentItem
appointment.Subject = "Report review"
appointment.Location = "Room 1"
appointment.Start = DateSerial(2026, 12, 1) + TimeSerial(10, 0, 0)
appointment.Duration = 45
WScript.Echo "Appointment: " & appointment.Subject & ", " & Year(appointment.Start) & "-" & Month(appointment.Start) & "-" & Day(appointment.Start) & _
    " " & Hour(appointment.Start) & ":" & Right("0" & Minute(appointment.Start), 2) & " - " & Hour(appointment.End) & ":" & Right("0" & Minute(appointment.End), 2) & _
    ", " & appointment.Location & ", reminder " & appointment.ReminderMinutesBeforeStart & " minutes before"

If dryRun Then
    WScript.Echo "Dry run: the mail is not shown"
Else
    mail.Display
End If
fso.DeleteFile file
