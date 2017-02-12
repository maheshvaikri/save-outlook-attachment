VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SaveAllAttachmentsInFolder()
    Dim folder As Outlook.folder
    
    Set folder = FindFolderModule.FindFolder(Configuration.SEARCH_FOLDER)
    If folder Is Nothing Then
        MsgBox "La carpeta '" & Configuration.SEARCH_FOLDER & "' no se encuentra"
        Exit Sub
    End If
    
    SaveAttachmentsInFolderItems.Save folder
End Sub
