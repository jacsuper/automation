
'Clear Deleted
Sub RemoveAllItemsAndFoldersInDeletedItems()
    Dim oDeletedItems As Outlook.folder
    Dim oFolders As Outlook.Folders
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    Set oItems = oDeletedItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.item(i).Delete
    Next
    Set oFolders = oDeletedItems.Folders
    For i = oFolders.Count To 1 Step -1
        oFolders.item(i).Delete
    Next
End Sub