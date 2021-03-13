'Duplicate current mail into a new mail x times
Public Sub DuplicateCurrentMail()
    Dim oArchive As Outlook.folder
    Dim oInspector As Inspector
    
    Dim olItem As Outlook.MailItem
    
    Dim windowType As String
    
    windowType = GetActiveSelectionType()
    
    If windowType = "List" Then
        If Application.ActiveExplorer.Selection.Count > 1 Then
            Dim Response As VbMsgBoxResult
            Response = MsgBox("Do you REALLY want to duplicate all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
            If Response = vbNo Then Exit Sub
        End If
        
        For Each olItem In Application.ActiveExplorer.Selection
          DoDuplicate olItem
        Next
        
    End If
    
    If windowType = "Email" Then
        Set oInspector = Application.ActiveInspector
        
        If Not oInspector Is Nothing Then
            Set olItem = oInspector.CurrentItem
            
            If Not olItem Is Nothing Then
                DoDuplicate olItem
            End If
        End If
        
    End If
End Sub


'Do Duplicate
Private Function DoDuplicate(olItem As Outlook.MailItem) As Boolean
    Dim newMsg As Outlook.MailItem

    Dim duplicateCount As String

    duplicateCount = InputBox("How many?", "", "1")

    If duplicateCount = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(duplicateCount) Then
        Exit Function
    End If
    
    Dim times As Integer
    times = duplicateCount
    
    For i = times To 1 Step -1

        Set newMsg = olItem.Forward 'Forward to make attachments that are hidden easy
    
        With newMsg
            .Subject = olItem.Subject 'Remove FWD:
            .HTMLBody = olItem.HTMLBody 'Remove change to Body from FWD
            .Display
        End With

    Next

    DoDuplicate = True
    
End Function
