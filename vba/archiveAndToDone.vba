'Do Done/Archive
Private Function DoDone(olItem As Object, oArchive As Outlook.folder) As Boolean

    olItem.UnRead = False
    olItem.Categories = "Done"
                
    If TypeOf olItem Is Outlook.MailItem Then
        If Not olItem.IsMarkedAsTask Then
            olItem.MarkAsTask olMarkComplete
        End If
    End If
        
    olItem.Save
        
    olItem.FlagStatus = olFlagComplete
        
    olItem.Save
        
    olItem.Move oArchive
 

    DoDone = True
End Function


'Active Window
Function GetActiveSelectionType() As String
    Dim objApp As Outlook.Application
           
    GetActiveSelectionType = ""
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            GetActiveSelectionType = "List"
        Case "Inspector"
            GetActiveSelectionType = "Email"
    End Select
       
    
End Function

'Archive email
Public Sub SimulateDoneQuickStep()
    Dim oArchive As Outlook.folder
    Dim oInspector As Inspector
    
    Dim olItem As Object
    
    'Obtain a reference to archive items folder
    Set oArchive = Session.GetDefaultFolder(olFolderInbox).Folders("!System").Folders("Archive")
    
    Dim windowType As String
    
    windowType = GetActiveSelectionType()
    
    If windowType = "List" Then
        If Application.ActiveExplorer.Selection.Count > 1 Then
            Dim Response As VbMsgBoxResult
            Response = MsgBox("Do you REALLY want to done all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
            If Response = vbNo Then Exit Sub
        End If
        
        For Each olItem In Application.ActiveExplorer.Selection
          DoDone olItem, oArchive
        Next
        
    End If
    
    If windowType = "Email" Then
        Set oInspector = Application.ActiveInspector
        
        If Not oInspector Is Nothing Then
            Set olItem = oInspector.CurrentItem
            
            If Not olItem Is Nothing Then
                DoDone olItem, oArchive
            End If
        End If
        
    End If

End Sub