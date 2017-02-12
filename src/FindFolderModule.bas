Attribute VB_Name = "FindFolderModule"
Function FindFolder(ByVal folderName As String) As Outlook.folder
 Dim colStores As Outlook.Stores
 Dim oStore As Outlook.Store
 Dim oRoot As Outlook.folder
 Dim found As Outlook.folder
 
 On Error Resume Next
 Set colStores = Application.Session.Stores
 
 For Each oStore In colStores
    Set oRoot = oStore.GetRootFolder
    ' Debug.Print (oRoot.FolderPath)
    Set found = FindFolderRecursively(oRoot, folderName)
    If Not found Is Nothing Then
        Set FindFolder = found
        Exit Function
    End If
 Next
End Function

Private Function FindFolderRecursively(ByVal oFolder As Outlook.folder, ByVal folderName As String) As Outlook.folder
    Dim folders As Outlook.folders
    Dim folder As Outlook.folder
    Dim found As Outlook.folder
    Dim foldercount As Integer
    
    On Error Resume Next
    Set folders = oFolder.folders
    foldercount = folders.count
    
    'Check if there are any folders below oFolder
    If foldercount = 0 Then
        FindFolderRercusively = Nothing
        Exit Function
    End If
        
    For Each folder In folders
        If folder.Name = folderName Then
            ' Debug.Print "Found: " & (folder.Name)
            Set found = folder
        Else
            Set found = FindFolderRecursively(folder, folderName)
        End If
    
        If Not found Is Nothing Then
            Set FindFolderRecursively = found
            Exit Function
        End If
    Next
End Function
