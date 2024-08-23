Attribute VB_Name = "LottoAttachments"
Sub Comdata_SaveAttachments(objMail As MailItem)
    
    Dim objAttachment As Attachment
    Dim strDownloadFolder As String
    Dim strFileType As String
    
    ' Set the download folder path
    strDownloadFolder = "C:\Users\smitchell\Desktop\Outlook Attachments\Comdata\"
    
    ' Set the file type to filter attachments
    strFileType = "xls"
    
    ' Loop through each attachment in the mail item
    For Each objAttachment In objMail.Attachments
        ' Check if the attachment's file type matches the specified file type
        If LCase(Right(objAttachment.FileName, Len(strFileType))) = strFileType Then
            ' Save the attachment to the download folder
            objAttachment.SaveAsFile strDownloadFolder & objAttachment.FileName
        End If
    Next objAttachment
    
    
    
End Sub

Sub Lottery_SaveAttachments(objMail As MailItem)
    Dim objAttachment As Attachment
    Dim strDownloadFolder As String
    Dim strFileType As String
    
    ' Set the download folder path
    strDownloadFolder = "C:\Users\smitchell\Desktop\Outlook Attachments\Lottery\"
    
    ' Set the file type to filter attachments
    strFileType = "csv"
    
    ' Loop through each attachment in the mail item
    For Each objAttachment In objMail.Attachments
        ' Check if the attachment's file type matches the specified file type
        If LCase(Right(objAttachment.FileName, Len(strFileType))) = strFileType Then
            ' Save the attachment to the download folder
            objAttachment.SaveAsFile strDownloadFolder & objAttachment.FileName
        End If
    Next objAttachment
End Sub




