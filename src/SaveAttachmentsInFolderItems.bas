Attribute VB_Name = "SaveAttachmentsInFolderItems"
Sub Save(ByVal folder As Outlook.folder)
    Dim i As Integer
    Dim item As Outlook.MailItem
    
    'Debug.Print "Found: " & folder.Items.count & " items in folder: " & folder.Name
    
    ' Iterar por todos los items de la carpeta de Outlook
    For i = 1 To folder.Items.count
        Set item = folder.Items(i)
        SaveAllAttachments item.attachments
    Next
    
End Sub

Private Sub SaveAllAttachments(ByVal attachments As Outlook.attachments)
    Dim i As Integer
    
    For i = 1 To attachments.count
        SaveAnAttachment attachments.item(i)
    Next
End Sub

Private Sub SaveAnAttachment(ByVal attachment As Outlook.attachment)
    'Debug.Print "Saving attachment: " & attachment.DisplayName & " to path " & Configuration.SAVE_PATH
    MsgBox "Guardando: " & attachment.DisplayName & " en el directorio " & Configuration.SAVE_PATH
    
    attachment.SaveAsFile (Configuration.SAVE_PATH & attachment.DisplayName)
End Sub

