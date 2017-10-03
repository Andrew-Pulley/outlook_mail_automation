Attribute VB_Name = "NISEmails"
'Define globalSettings types
Type settings
 firstName As String
 lastName As String
 numberExtension As String
End Type
'Define globalSettings values
Function getSettings() As settings
    Dim sets As settings

    With sets 'Retreive values
.firstName = "First"
.lastName = "Last"
.numberExtension = "1234"
    End With
    'Return a single struct
    getSettings = sets
End Function

Sub CalendarInvite()
'Form to send calendar invite.
formAppointment.Show
End Sub

Sub ContactInformation()
'Contact Information Template
'Define variables
Dim emailPath As String
Dim emailSubject As String
emailPath = "P:\Departments\NIS\Outbound\Email Templates\contactInformation.html"
emailSubject = "Thank you for your time!"
Call generateEmail(emailPath, emailSubject, "")
End Sub

Sub MissedCall()
'Missed Call Template
'Define variables
Dim emailPath As String
Dim emailSubject As String

emailPath = "P:\Departments\NIS\Outbound\Email Templates\missedCall.html"
emailSubject = "Sorry I Missed You!"

Call generateEmail(emailPath, emailSubject, "")
End Sub

Sub SATLead()
'SAT Lead Follow Up Template

'Load from sheet


'Define email variables
Dim emailPath As String
Dim emailSubject As String
Dim customerName As String

emailPath = "P:\Departments\NIS\Outbound\Email Templates\satLead.html"
emailSubject = "Just following up."
customerName = "Test Customer"

Call generateEmail(emailPath, emailSubject, customerName)
End Sub

Sub Quote()
'Quote Template
'Define variables
Dim emailPath As String
Dim emailSubject As String

emailPath = "P:\Departments\NIS\Outbound\Email Templates\quote.html"
emailSubject = "Your vivint.SmartHome Quote"

Call generateEmail(emailPath, emailSubject, "")
End Sub

Function generateEmail(emailPath As String, emailSubject As String, customerName As String)
'Retrieve globalSettings
 Dim globalSettings As settings
 globalSettings = getSettings()

'Copy email address from clipboard
Dim DataObj As MSForms.DataObject
 Set DataObj = New MSForms.DataObject
 DataObj.GetFromClipboard
strPaste = DataObj.GetText(1)

'Define variablized email link
Dim emailLink As String
 emailLink = "<a href='mailto:" _
 & globalSettings.firstName _
 & "." _
 & globalSettings.lastName _
 & "@vivint.com?utm_source=outbound_lead" _
 & "&utm_medium=email&utm_campaign=missed_call_email'" _
 & "style='width:250; display:block; text-decoration:none;" _
 & "border:0; text-align:center: font-weight:bold;" _
 & "font-size:18px; font-family: Arial, sans-serif;" _
 & "color: #ffffff' class='button_link'>Schedule a Call</a>"
 
'Create email w/ populated "To" and "Subject" fields
Dim MItem As MailItem
Set MItem = Application.CreateItem(olMailItem)
MItem.To = strPaste
MItem.Subject = emailSubject
MItem.Body = ""
MItem.Display

'Populate Body
Dim insp As Inspector
Set insp = ActiveInspector

'Render HTML with Word Renderer
If insp.IsWordMail Then
Dim wordDoc As Word.Document
Set wordDoc = insp.WordEditor
wordDoc.Application.Selection.InsertFile _
 emailPath _
 , , False, False, False
End If

'Replace variable text with globalVariables
MItem.HTMLBody = Replace(MItem.HTMLBody, "%CustomerName%", customerName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%FirstName%", globalSettings.firstName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%LastName%", globalSettings.lastName)
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%NumberExtension%", globalSettings.numberExtension)
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%EmailLink%", emailLink)
End Function

Function generateScheduledEmail(emailPath As String, emailSubject As String, customerName As String)
'Retrieve globalSettings
 Dim globalSettings As settings
 globalSettings = getSettings()

'Copy email address from clipboard
Dim DataObj As MSForms.DataObject
 Set DataObj = New MSForms.DataObject
 DataObj.GetFromClipboard
strPaste = DataObj.GetText(1)

'Define variablized email link
Dim emailLink As String
 emailLink = "<a href='mailto:" _
 & globalSettings.firstName _
 & "." _
 & globalSettings.lastName _
 & "@vivint.com?utm_source=outbound_lead" _
 & "&utm_medium=email&utm_campaign=missed_call_email'" _
 & "style='width:250; display:block; text-decoration:none;" _
 & "border:0; text-align:center: font-weight:bold;" _
 & "font-size:18px; font-family: Arial, sans-serif;" _
 & "color: #ffffff' class='button_link'>Schedule a Call</a>"
 
'Create email w/ populated "To" and "Subject" fields
Dim MItem As MailItem
Set MItem = Application.CreateItem(olMailItem)
MItem.To = strPaste
MItem.Subject = emailSubject
MItem.DeferredDeliveryTime = sendDate
MItem.Body = ""
MItem.Display

'Populate Body
Dim insp As Inspector
Set insp = ActiveInspector

'Render HTML with Word Renderer
If insp.IsWordMail Then
Dim wordDoc As Word.Document
Set wordDoc = insp.WordEditor
wordDoc.Application.Selection.InsertFile _
 emailPath _
 , , False, False, False
End If

'Replace variable text with globalVariables
MItem.HTMLBody = Replace(MItem.HTMLBody, "%CustomerName%", customerName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%FirstName%", globalSettings.firstName)
MItem.HTMLBody = Replace(MItem.HTMLBody, "%LastName%", globalSettings.lastName)
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%NumberExtension%", globalSettings.numberExtension)
MItem.HTMLBody = Replace _
 (MItem.HTMLBody, "%EmailLink%", emailLink)
End Function

