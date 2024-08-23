Attribute VB_Name = "Email"
Sub EmailDevin()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim FilePath As String
    
    ' Create a new instance of Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Define the file path of the workbook
    FilePath = ActiveWorkbook.FullName
    
    ' Add the workbook as an attachment
    OutlookMail.Attachments.Add FilePath
    
    ' You can customize the email subject and body here
    OutlookMail.Subject = Replace(ActiveWorkbook.Name, Right(ActiveWorkbook.Name, 5), "")
    OutlookMail.Body = ""
    
    ' Add the recipient's email address
    OutlookMail.to = "dcharles@foodnfun.com"
'    OutlookMail.To = "smitchell@foodnfun.com"
    
    ' Optionally, you can add CC and BCC recipients
    ' OutlookMail.CC = "cc@example.com"
    ' OutlookMail.BCC = "bcc@example.com"
    
    ' Send the email
    OutlookMail.Send
    
    ' Release the Outlook objects
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    ActiveWorkbook.Close
End Sub

Sub EmailTony()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim FilePath As String
    
    ActiveWorkbook.Save
    
    ' Create a new instance of Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Define the file path of the workbook
    FilePath = ActiveWorkbook.FullName
    
    ' Add the workbook as an attachment
    OutlookMail.Attachments.Add FilePath
    
    ' You can customize the email subject and body here
    OutlookMail.Subject = Replace(ActiveWorkbook.Name, Right(ActiveWorkbook.Name, 5), "")
    OutlookMail.Body = ""
    
    ' Add the recipient's email address
    OutlookMail.to = "tridge@foodnfun.com"
    
    ' Optionally, you can add CC and BCC recipients
    ' OutlookMail.CC = "cc@example.com"
    ' OutlookMail.BCC = "bcc@example.com"
    
    ' Send the email
    OutlookMail.Send
    
    ' Release the Outlook objects
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

