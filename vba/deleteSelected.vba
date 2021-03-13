'Macro to move to deleted items after setting category to nothing
Public Sub DeleteSelected()
    Dim oDeletedItems As Outlook.folder
    
    Dim olItem As Outlook.MailItem
    
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Do you REALLY want to delete all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
        If Response = vbNo Then Exit Sub
    End If
    
    For Each olItem In Application.ActiveExplorer.Selection

        olItem.UnRead = False
        olItem.Categories = ""
        olItem.Save
        
        olItem.Move oDeletedItems
 
    Next
  

End Sub