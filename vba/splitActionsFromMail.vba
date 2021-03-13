
'Split out new mails for all * action items in notes
'Add Microsoft VB Script Regular Expressions 5.5 Reference
Public Sub SplitOutActionMails()
     Dim olItem As Outlook.MailItem
     
     Dim reg1 As RegExp
     Dim M1 As MatchCollection
     Dim M As match
     
     Dim objMsg As MailItem
     
     Dim str As String
    
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Do you REALLY want to Action all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
        If Response = vbNo Then Exit Sub
    End If
    
    Set reg1 = New RegExp
    
    For Each olItem In Application.ActiveExplorer.Selection

        olItem.UnRead = False
        
        olItem.Save
        
       
         With reg1
           .Pattern = "\*[a-zA-Z ]+\n*"
           .Global = True
         End With
         If reg1.Test(olItem.Body) Then
    
            Set M1 = reg1.Execute(olItem.Body)
            For Each M In M1
                str = M.Value
                Debug.Print str
                
                Set objMsg = Application.CreateItem(olMailItem)
    
                With objMsg
                  .To = "email@email.com"
                  .CC = ""
                  .BCC = ""
                  .Subject = StripFolderNameOrSubject(olItem.Subject) & ": " & str
                  .BodyFormat = olFormatHTML
                  .Body = ""
                    
                  .Send
                
                
                End With
            Next
          End If

 
    Next
End Sub
