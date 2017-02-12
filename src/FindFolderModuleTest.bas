Attribute VB_Name = "FindFolderModuleTEst"
Sub TEST_FindFolder_Finds_A_Known_Folder()
    Dim folder As Outlook.folder
    Dim searchFolderName As String
    
    searchFolderName = "Test Macro"
    Set folder = FindFolderModule.FindFolder(searchFolderName)
    
    Debug.Assert Not folder Is Nothing
    If folder Is Nothing Then
        Debug.Print "Folder " & searchFolderName & " should have been found"
    End If
End Sub

Sub TEST_FindFolder_Cant_Find_An_Unknown_Folder()
    Dim folder As Outlook.folder
    Dim searchFolderName As String
    
    searchFolderName = "Test Macro 2"
    Set folder = FindFolderModule.FindFolder(searchFolderName)
    
    Debug.Assert folder Is Nothing
    If Not folder Is Nothing Then
        Debug.Print "Folder " & searchFolderName & " should NOT have been found"
    End If
End Sub

