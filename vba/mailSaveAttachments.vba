Private WithEvents Items As Outlook.Items
Private WithEvents MailItems As Outlook.Items


Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"


'Bind some events on startup
Private Sub Application_Startup()
    Dim folder As Outlook.folder
    
    'Bind new Inbox items
    Set MailItems = Session.GetDefaultFolder(olFolderInbox).Items
    
End Sub

'Mail arrived in Inbox - save the Attachment to today's attachment folder
Private Sub MailItems_ItemAdd(ByVal item As Object)
    On Error GoTo MailItems_ItemAdd_Error

    Dim objAtt As Outlook.Attachment
    Dim saveFile As String
    Dim dateFormat As String
    
    If TypeOf item Is Outlook.AppointmentItem Or TypeOf item Is Outlook.MailItem Then
    
        dateFormat = Format(item.ReceivedTime, "yyyymmdd")
        Dim saveFolder As String
        saveFolder = "C:\Users\" & USERNAME & "\Desktop\Attachments" & "\" & dateFormat
        If (Dir$(saveFolder, vbDirectory) = "") Then
            MkDir saveFolder
        End If
        For Each objAtt In item.Attachments
            If objAtt.Type = olByValue Then
                saveFile = saveFolder & "\" & objAtt.DisplayName
                objAtt.SaveAsFile saveFile
                'MsgBox ("Saved " & saveFile)
            End If
        Next
        
    End If
    
    Exit Sub
        
MailItems_ItemAdd_Error:

    Exit Sub
    
End Sub