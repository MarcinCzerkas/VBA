Sub SendEmailWithTable()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim rng As Range
    Dim tableHTML As String
    
    ' Set the worksheet and range containing the table
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change as needed
    Set rng = ws.Range("A1:D5") ' Change to your table range

    ' Convert range to HTML table
    tableHTML = RangeToHTML(rng)
    
    ' Create Outlook Application
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Prepare email
    With OutlookMail
        .To = "recipient@example.com"
        .CC = ""
        .BCC = ""
        .Subject = "Excel Table in Email"
        .HTMLBody = "Dear Recipient,<br><br>Please find the table below:<br><br>" & tableHTML & "<br><br>Best regards,<br>Your Name"
        .Display ' Use .Send to send the email automatically
    End With
    
    ' Cleanup
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

Function RangeToHTML(rng As Range) As String
    Dim TempWorkbook As Workbook
    Dim TempWorksheet As Worksheet
    Dim TempRange As Range
    Dim HTMLString As String
    
    ' Create a temporary workbook
    Set TempWorkbook = Workbooks.Add
    Set TempWorksheet = TempWorkbook.Sheets(1)
    
    ' Copy the range to the temporary worksheet
    rng.Copy
    TempWorksheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    TempWorksheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    
    ' Save the range as an HTML file
    TempWorkbook.PublishObjects.Add(xlSourceRange, "C:\Temp\Table.html", TempWorksheet.Name, TempWorksheet.UsedRange.Address, xlHtmlStatic).Publish True
    
    ' Read the HTML file as a string
    Dim FileNum As Integer
    Dim FileContent As String
    FileNum = FreeFile
    Open "C:\Temp\Table.html" For Input As #FileNum
    FileContent = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    
    ' Clean up the temp workbook
    TempWorkbook.Close False
    Kill "C:\Temp\Table.html"
    
    ' Return the HTML table
    RangeToHTML = FileContent
End Function