Sub SendAllReports_MarcinCzerkas_11032024()

'[0] DOCUMENTATION
'This code (together with this spreadsheet) manages automatic sending of the weekly updates.

'HOW DOES THE CODE WORK? - SHORT EXPLATANTION
'The code loops through a list of e-mails and their attributes (recipients, cc, subject, body, attachment), finds the respective attachment, opens it, refreshes the queries, saves it and sends a new e-mail with the refreshed file attached.
'PREREQUISITES: a separate folder called "Countries" must be stored in the same directory as this file. The macro will automatically detect its path. The attachments should be stored in this folder. The files' names should be the same as the values in the column "A" in the sheet [List].

'Please follow the instructions provided in the sheet [Dashboard] before running the macro.

'[1] PREPARATION STEPS

'Declare the variables
Dim ThisFilePath As String
Dim SureQuestion As String
Dim AttachmentsPath As Variant

'Define the folder path
ThisFilePath = ThisWorkbook.Path
AttachmentsPath = (ThisFilePath & "\Countries\")
    
'Are-you-sure-Question
SureQuestion = MsgBox("Do you want to start sending emails?" & Chr(13) & Chr(13) & "Important: All previous steps must be done", vbYesNo + vbQuestion, "Report Sending")
If SureQuestion = vbNo Then
    MsgBox "Process cancelled", vbOKOnly, "Report Sending"
    Exit Sub
    Else
        
    'Turn off unnecessary stuff to make the macro run faster
    Calculate
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

'[2] REFRESH AND SEND THE REPORTS

    'Declare the variables
    Dim OutlookApp As Outlook.Application
    Dim MItem As Outlook.MailItem
    Dim LastMail As String
    Dim MailTo As String
    Dim MailCC As String
    Dim strFileExists As String
    Dim wb As Workbook

    Set OutlookApp = New Outlook.Application

    'Check the number of e-mails to be sent
    LastMail = ThisWorkbook.Sheets("List").Range("A1").CurrentRegion.Rows.Count

    'Creating the emails (using For Next Loop)
    For Each cell In ThisWorkbook.Sheets("List").Range("A2:A" & LastMail)

        'Declare variables for specific parts of the email
        MailTo = cell.Offset(0, 3).Value    'Recipients
        MailSubject = cell.Offset(0, 5).Value   'Subject
        MailCC = cell.Offset(0, 4).Value    'CC
        MailAttachment = (AttachmentsPath & cell.Offset(0, 1).Value & "*.xlsx") 'Attachment

        'Check if the attachment exists
        strFileExists = Dir(MailAttachment)
        
        ' Check if three conditions are met: 1) the attachment exists, 2) the selected row (recipients) is marked as "YES" which means this e-mail is supposed to be sent, 3) the value in the column "I" =1 which indicates that today's day of the week corresponds to the day of week when the e-mail should be sent
        If strFileExists <> "" And cell.Offset(0, 2).Value = "YES" And cell.Offset(0, 8).Value = 1 Then
            
            'Refresh the file that we want to attach
            Set wb = Workbooks.Open(MailAttachment)
            
            'Turn on FastCombine to skip the pop-up and speed up the refresh
            wb.Queries.FastCombine = True
            
            'Refresh the Power Query connections for all queries
            wb.Connections("Query - Special Requests").Refresh
            wb.Connections("Query - Open Requests").Refresh
            wb.Connections("Query - Closed Requests").Refresh
            wb.Connections("Query - Clarifications").Refresh
            
            'Save the current date in the homepage of the attachment
            wb.Sheets("Legend").Range("D3").Value = Date
            
            'Save the report and update the date in its name
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            Name AttachmentsPath & strFileExists As AttachmentsPath & cell.Offset(0, 1).Value & " " & Date & ".xlsx"
            
            'Send the email
            Calculate
            Set MItem = OutlookApp.CreateItem(olMailItem)
            With MItem
                .SentOnBehalfOfName = ThisWorkbook.Sheets("Dashboard").Range("D13").Value   'Send from the team shared mailbox indicated in the sheet [Dashboard]
                .To = MailTo    'Recipient
                If MailCC <> "" Then .cc = MailCC   'CC
                .Subject = MailSubject  'Subject
                .Body = cell.Offset(0, 6).Value 'E-mail body
                .Attachments.Add AttachmentsPath & cell.Offset(0, 1).Value & " " & Date & ".xlsx"   'Attachment
                .Send
            End With

            Else
        End If
    Next cell

End If


'[3] FINAL STATEMENTS AND CLEARING

'Update the journal with the date and time of the refresh as well with the username
Dim LastJournalEntry As Integer
LastJournalEntry = Journal.Cells(Rows.Count, "A").End(xlUp).Row + 1
Journal.Cells(LastJournalEntry, 1).Value = Date
Journal.Cells(LastJournalEntry, 2).Value = Time
Journal.Cells(LastJournalEntry, 3).Value = Application.UserName

'Clean up
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.ScreenUpdating = True

'Message box informing about the execution of the macro
MsgBox "Emails Sent" & Chr(13) & Chr(13) & "Process finilized successfully", vbInformation, "Report Sending"

'{!} This code still needs an error handling section {!}

End Sub