
' Copy attachments
Sub CopyAttachments(objSourceItem, objTargetItem)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = fso.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp.Path & "\"
   For Each objAtt In objSourceItem.Attachments
      strFile = strPath & objAtt.Filename
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, olByValue, , objAtt.DisplayName
      fso.DeleteFile strFile
   Next

   Set fldTemp = Nothing
   Set fso = Nothing
End Sub

'Reply and keep attachments
Sub ReplyWithAttachments()
    Dim oReply As Outlook.MailItem
    Dim oItem As Object
     
    Set oItem = GetCurrentItem()
    If Not oItem Is Nothing Then
        Set oReply = oItem.Reply
        CopyAttachments oItem, oReply
        oReply.Display
        oItem.UnRead = False
    End If
     
    Set oReply = Nothing
    Set oItem = Nothing
End Sub

'Reply all and keep attachments
Sub ReplyAllWithAttachments()
    Dim oReply As Outlook.MailItem
    Dim oItem As Object
     
    Set oItem = GetCurrentItem()
    If Not oItem Is Nothing Then
        Set oReply = oItem.ReplyAll
        CopyAttachments oItem, oReply
        oReply.Display
        oItem.UnRead = False
    End If
     
    Set oReply = Nothing
    Set oItem = Nothing
End Sub
