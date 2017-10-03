VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAppointment 
   Caption         =   "Appointment Setter"
   ClientHeight    =   8715.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5400
   OleObjectBlob   =   "formAppointment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buttonOk_Click()

'Define variables
Dim emailLink As String
Dim startTime As String

'Validated, hide
formAppointment.Hide

'Define timezone offset
Dim startTimeInZone As String
Dim endTimeInZone As String
startTimeInZone = CStr(CInt(comboStartTime.Column(1)) _
 + CInt(comboTimeZone.Column(1))) & ":00"
endTimeInZone = CStr(CInt(comboEndTime.Column(1)) _
 + CInt(comboTimeZone.Column(1))) & ":00"

'Set appointment start, end, and timezone
calendarStartTime = comboMonth.Column(1) & "/" & textDay & "/" _
 & textYear & " " & startTimeInZone

calendarEndTime = comboMonth.Column(1) & "/" & textDay & "/" _
 & textYear & " " & endTimeInZone
 
emailStartTime = comboMonth.Column(1) & "/" & textDay & "/" _
 & textYear & " " & comboStartTime
emailEndTime = comboMonth.Column(1) & "/" & textDay & "/" _
 & textYear & " " & comboEndTime
 
calendarMinusOne = comboMonth.Column(1) & "/" & CStr(CInt(textDay) - 1) & "/" _
 & textYear & " " & startTimeInZone

Debug.Print (VarType(calendarMinusOne))

'Create appointment
Dim AItem As AppointmentItem
Set AItem = Application.CreateItem(olAppointmentItem)

'Set meeting parameters
AItem.MeetingStatus = olMeeting
AItem.Subject = "Your vivint.SmartHome Installation - " _
 & textOpportunityNumber
AItem.OptionalAttendees = textCustomerEmail
AItem.Location = textLocation
AItem.Start = CDate(calendarStartTime)
AItem.End = CDate(calendarEndTime)
AItem.Display

'Create email
Dim MItem As MailItem
Set MItem = Application.CreateItem(olMailItem)
MItem.To = textCustomerEmail
MItem.Subject = "Welcome to Vivint!"
MItem.Body = ""
MItem.Display

'Render HTML template with Word
Dim insp As Inspector
Set insp = ActiveInspector
If insp.IsWordMail Then
Dim wordDoc As Word.Document
Set wordDoc = insp.WordEditor
wordDoc.Application.Selection.InsertFile _
 "P:\Departments\NIS\Outbound\Email Templates\calendarInvite.html" _
 , , False, False, False
End If

'Replace variables in HTMLBody
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%NumberExtension%", labelNumberExtension)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%CustomerName%", textCustomerName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%CustomMessage%", textCustomMessage)
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%CalendarDate%", Format(emailStartTime, "dddd, mmmm dd,"))
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%StartTime%", Format(emailStartTime, "h am/pm"))
MItem.HTMLBody = Replace _
(MItem.HTMLBody, "%EndTime%", Format(emailEndTime, "h am/pm"))
MItem.HTMLBody = Replace(MItem.HTMLBody, "%CustomMessage%", textCustomMessage)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%FirstName%", labelFirstName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%LastName%", labelLastName)

'Quote Template
'Define variables
'Dim emailPath As String
'Dim emailSubject As String

'emailPath = "P:\Departments\NIS\Outbound\Email Templates\preInstall.html"
'emailSubject = "What to expect at your installation."
'sendDate = CDate(calendarMinusOne)
'Call generateScheduledEmail(emailPath, emailSubject, textCustomerName)
End Sub


Private Sub UserForm_Initialize()
 'Define variables
 Dim firstName As String
 Dim lastName As String
 Dim numberExtension As String
 Dim itemsMonth() As String
 Dim itemsMonthValue() As String
 Dim i As Long
 Dim itemsTime() As String
 Dim itemsTimeValue() As String
 Dim itemsTimeZone() As String
 Dim itemsTimeZoneValue() As String

 'Get global settings
 Dim globalSettings As settings
 globalSettings = getSettings()
 labelFirstName = globalSettings.firstName
 labelLastName = globalSettings.lastName
 labelNumberExtension = globalSettings.numberExtension
 
 'Copy email address from clipboard
 Dim DataObj As MSForms.DataObject
 Set DataObj = New MSForms.DataObject
 DataObj.GetFromClipboard
 textCustomerEmail = DataObj.GetText(1)
 
 'Populate user-friendly month list
 itemsMonth = Split("January|February|March|April|May|" _
  & "June|July|August|September|October|November|December", "|")
 itemsMonthValue = Split("01|02|03|04|05|06|07|08|09|10|11|12", "|")
 comboMonth.ColumnWidths = "60;0"
 
 For i = 0 To UBound(itemsMonth)
  comboMonth.AddItem
  comboMonth.List(i, 0) = itemsMonth(i)
  comboMonth.List(i, 1) = itemsMonthValue(i)
 Next i
 
 'Populate times
 itemsTime = Split("8:00AM|9:00AM|10:00AM|" _
  & "11:00AM|12:00PM|1:00PM|2:00PM|3:00PM|" _
  & "4:00PM|5:00PM|6:00PM|7:00PM|8:00PM", "|")
 itemsTimeValue = Split("8|9|10|11|12|13|14|15|16|17|18|19|20", "|")
 
 comboStartTime.ColumnWidths = "60;0"
 comboEndTime.ColumnWidths = "60;0"
 
 For i = 0 To UBound(itemsTime)
  comboStartTime.AddItem
  comboStartTime.List(i, 0) = itemsTime(i)
  comboStartTime.List(i, 1) = itemsTimeValue(i)
  comboEndTime.AddItem
  comboEndTime.List(i, 0) = itemsTime(i)
  comboEndTime.List(i, 1) = itemsTimeValue(i)
 Next i
  
 'Populate user-friendly timezones
 itemsTimeZone = Split("Hawaii|" _
  & "Alaska|Pacific Time|" _
  & "Arizona|Mountain Time|" _
  & "Central Time|Eastern Time", "|")
 itemsTimeZoneValue = Split("3|2|1|1|0|-1|-2", "|")
 comboTimeZone.ColumnWidths = "60;0"
 
 For i = 0 To UBound(itemsTimeZone)
  comboTimeZone.AddItem
  comboTimeZone.List(i, 0) = itemsTimeZone(i)
  comboTimeZone.List(i, 1) = itemsTimeZoneValue(i)
 Next i

 'Populate custom message
  textCustomMessage = _
   "Thank you so much for your time today!" _
   & " You'll love your new Vivint system."
End Sub
