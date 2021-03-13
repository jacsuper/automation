Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"

'Move all in folder to Deleted Items
Private Function DeleteFolderContents(folderPath As String) As Boolean
    Dim oDeletedItems As Outlook.folder
    Dim oFolder As Outlook.folder
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    Set oFolder = GetFolderPath(folderPath)
    
    If Not oFolder Is Nothing Then
        Set oItems = oFolder.Items
        For i = oItems.Count To 1 Step -1
            oItems.item(i).Move oDeletedItems
        Next
        DeleteFolderContents = True
        Exit Function
    End If
    
    DeleteFolderContents = False

  
End Function
'Mark all in folder as read
Private Function MarkAsReadFolderContents(folderPath As String) As Boolean
    Dim oFolder As Outlook.folder
    Dim oItems As Outlook.Items
    Dim i As Long

    Set oFolder = GetFolderPath(folderPath)
    
    If Not oFolder Is Nothing Then
        Set oItems = oFolder.Items
        For i = oItems.Count To 1 Step -1
            If (oItems.item(i).UnRead = True) Then
                oItems.item(i).UnRead = False
            End If
        Next
        MarkAsReadFolderContents = True
        Exit Function
    End If
    
    MarkAsReadFolderContents = False

  
End Function

'Move all in my junk folders to Deleted
Public Sub ClearJunkFolders()
    DeleteFolderContents ("\\" & USERNAME & "\Inbox\!System\KB\KPI")

End Sub
