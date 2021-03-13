Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"


'Move out attachments from selected to one drive
Sub Delete_attachments_SaveToOneDrive()

 Dim myAttachment As Attachment
 Dim myAttachments As Attachments
 Dim selItems As Selection
 Dim myItem As Object
 Dim lngAttachmentCount As Long
 Dim newFileName As String
 
 ' Set reference to the Selection.
 Set selItems = ActiveExplorer.Selection

 If selItems.Count > 1 Then
    Dim Response As VbMsgBoxResult
    Response = MsgBox("Do you REALLY want to delete all attachments in all SELECTED mails?" _
    , vbExclamation + vbDefaultButton2 + vbYesNo)
    If Response = vbNo Then Exit Sub
 End If

 ' Loop though each item in the selection.
 For Each myItem In selItems
     Set myAttachments = myItem.Attachments
    
     lngAttachmentCount = myAttachments.Count
    
     ' Loop through attachments until attachment count = 0.
     While lngAttachmentCount > 0

        newFileName = myAttachments.item(1).Filename
        
        newFileName = "C:\Users\" & USERNAME & "\Attachments\" & newFileName
 
        myAttachments(1).SaveAsFile newFileName
        
        strFile = newFileName & "; " & strFile
        
        myAttachments(1).Delete
        lngAttachmentCount = myAttachments.Count
     Wend
     
     
     If myItem.BodyFormat <> olFormatHTML Then
        myItem.Body = myItem.Body & vbCrLf & _
            "The file(s) removed were: " & strFile
     Else
        myItem.HTMLBody = myItem.HTMLBody & "<p>" & _
            "The file(s) removed were: " & strFile & "</p>"
     End If
     
    
     myItem.Save
      strFile = ""
 Next


    Set myAttachment = Nothing
    Set myAttachments = Nothing
    Set selItems = Nothing
    Set myItem = Nothing
    
 End Sub
