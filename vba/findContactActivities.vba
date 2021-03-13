'Find related mails todo with this contact from selected and show in new window
Sub FindContactActivities()

' =================================================================
' Description: Outlook macro to find all mails related to a contact
' author : Victor Beekman victor[dot]beekman"at"xs4all{dot}nl
' version: 1.1 160623
' Note: the "to:" and "from:" should be localized in your own language.
' Please check the proper wording by using the instant search feature.
'==================================================================

    If Application.Session.DefaultStore.IsInstantSearchEnabled Then
        Dim olkFilter As String
        Dim oItem As Object
        Dim sender As String
        Dim olkExplorer As Outlook.Explorer
        If Application.ActiveExplorer.Selection.Count <> 1 Then
            MsgBox "Command works only for one selected item", vbExclamation
            Exit Sub
        End If
        Set oItem = Application.ActiveExplorer.Selection.item(1)
        Set olkExplorer = Application.Explorers.Add(Application.Session.GetDefaultFolder(olFolderInbox), olFolderDisplayNormal)
        If oItem.Class = olContact Then
              olkFilter = "from:(" & Chr(34) & oItem.Email1Address & Chr(34) & ") OR " _
                 & " to:(" & Chr(34) & oItem.Email1Address & Chr(34) & ") OR " _
                 & "cc:(" & Chr(34) & oItem.Email1Address & Chr(34) & ")"
            If Len(oItem.Email2Address) > 5 Then
              olkFilter = olkFilter & " OR from:(" & Chr(34) & oItem.Email2Address & Chr(34) & ") OR " _
                 & " to:(" & Chr(34) & oItem.Email2Address & Chr(34) & ") OR " _
                 & "cc:(" & Chr(34) & oItem.Email2Address & Chr(34) & ")"
            End If
 
          ElseIf oItem.Class = olMail Then
            sender = oItem.SenderEmailAddress
            If StartsWith(sender, "/") Then
                sender = oItem.SenderName
            End If
            
          
            olkFilter = " from:(" & Chr(34) & sender & Chr(34) & ") OR " _
                 & "to:(" & Chr(34) & sender & Chr(34) & ") OR " _
                 & "cc:(" & Chr(34) & sender & Chr(34) & ")"
           Else
            MsgBox "Command can't be run from here", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Search Indexing is not enabled for this mailbox." & vbNewLine & vbNewLine & _
                    "If you are using an Exchange account, make sure that Cached Exchange Mode is enabled" & _
                    vbExclamation, "Search Indexing not enabled"
        Exit Sub
    End If
    Call olkExplorer.Search(olkFilter, olSearchScopeAllOutlookItems)
    Call olkExplorer.Display
' hide the navigation bar & preview pane
    olkExplorer.ShowPane olFolderList, False
    olkExplorer.ShowPane olOutlookBar, False
    olkExplorer.ShowPane olPreview, False
    olkExplorer.ShowPane olNavigationPane, False
    'olkExplorer.ShowPane olToDoBar, False
End Sub
