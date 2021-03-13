
'Make current mail confidential
Public Sub MakeConfidential()
    Dim olItem As Outlook.MailItem
    
    
    If Application.ActiveInspector Is Nothing Then
        Exit Sub
    End If
    
    Set olItem = Application.ActiveInspector.CurrentItem
    
    If Not olItem.Sent Then
        If (InStr(olItem.Subject, "[Confidential]") = 0) Then
            olItem.Subject = "[Confidential] " & olItem.Subject
        End If
        olItem.Sensitivity = olConfidential
        olItem.Save
    End If
End Sub


