
'Clear categories and flag
Public Sub ClearCategoriesAndFlag()
    Dim olItem As Outlook.MailItem
    
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Do you REALLY want to categorize all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
        If Response = vbNo Then Exit Sub
    End If
    
    For Each olItem In Application.ActiveExplorer.Selection

        olItem.UnRead = False
        olItem.Categories = ""
        olItem.Save
        
        If Not olItem.IsMarkedAsTask Then
            olItem.MarkAsTask olMarkComplete
        End If
            
        olItem.Save
            
        olItem.FlagStatus = olFlagComplete
        
        olItem.Save
 
    Next
End Sub