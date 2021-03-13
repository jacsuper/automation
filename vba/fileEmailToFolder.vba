Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"

'Normalize folder or subject
Private Function StripFolderNameOrSubject(name As String) As String
    StripFolderNameOrSubject = name
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, " Notes", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "1:1", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "FWD:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "fwd:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "Fwd:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "RE:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "re:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "Re:", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "!", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "?", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, ".", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, ":", " ")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "[", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "(", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "]", " ")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, ")", " ")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "*", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "%", "")
    StripFolderNameOrSubject = Replace(StripFolderNameOrSubject, "  ", " ")
End Function

'Equivalent meeting notes
Private Function IsFolderAndSubjectEquivalent(strSubject As String, strFolderName As String) As Boolean
    IsFolderAndSubjectEquivalent = StrComp(strFolderName, strSubject, vbTextCompare) = 0 Or _
                                    StartsWith(strFolderName, strSubject) Or _
                                    StartsWith(strSubject, strFolderName) Or _
                                    EndsWith(strFolderName, strSubject) Or _
                                    EndsWith(strSubject, strFolderName)
End Function

'Related Project
Private Function IsFolderAndSubjectRelated(strSubject As String, strFolderName As String) As Boolean
    Dim subjectWords() As String
    Dim folderNames() As String

    subjectWords() = Split(strSubject, " ")
    folderNames() = Split(strFolderName, " ")
    
    IsFolderAndSubjectRelated = False
    
    For i = LBound(subjectWords) To UBound(subjectWords)
        For j = LBound(folderNames) To UBound(folderNames)
            If (subjectWords(i) <> "" And subjectWords(i) <> "-") Then
                If StrComp(subjectWords(i), folderNames(j), vbTextCompare) = 0 Then
                
                    IsFolderAndSubjectRelated = True
                    Exit Function
                
                End If
            End If
        Next j
    
    Next i
    
End Function

'In a low tech way try and detect where to file an email to.  Start with meeting notes, then fall through to project folders.
'If no luck - suggest creating a new meeting or project folder and move
Public Sub LowTechFiler()
    Dim olItem As Outlook.MailItem
    Dim oFolder As Outlook.folder
    Dim newFolder As Outlook.folder
    
    Dim oFolders As Outlook.Folders
    
    Dim strSubject As String
    Dim strFolderName As String
    Dim moved As Boolean
    Dim newName As String
    
    For Each olItem In Application.ActiveExplorer.Selection
        moved = False
        strSubject = olItem.Subject
    
        strSubject = StripFolderNameOrSubject(strSubject)
        
        Set oFolder = GetFolderPath("\\" & USERNAME & "\Inbox\!System\Projects\!Meetings")
        
        Set oFolders = oFolder.Folders
            
        For i = oFolders.Count To 1 Step -1
                If Not moved Then
                  strFolderName = StripFolderNameOrSubject(oFolders.item(i).name)
                  If IsFolderAndSubjectEquivalent(strSubject, strFolderName) Then
                     
                      If MsgBox("Do you want to File to Meetings: " & oFolders.item(i).name, vbYesNo) = vbYes Then
                          olItem.UnRead = False
                          
                          olItem.Move oFolders.item(i)
                
                           moved = True
                      End If
                  End If
                End If
        Next
            
        If Not moved Then
            Set oFolder = GetFolderPath("\\" & USERNAME & "\Inbox\!System\Projects")
            
            Set oFolders = oFolder.Folders
            
            For i = oFolders.Count To 1 Step -1
                If Not moved Then
                  strFolderName = StripFolderNameOrSubject(oFolders.item(i).name)
                  If IsFolderAndSubjectEquivalent(strSubject, strFolderName) Or IsFolderAndSubjectRelated(strSubject, strFolderName) Then
                     
                      If MsgBox("Do you want to File to Projects: " & oFolders.item(i).name, vbYesNo) = vbYes Then
                          olItem.UnRead = False
                          
                          olItem.Move oFolders.item(i)
                
                           moved = True
                      End If
                  End If
                End If
            Next
            
            If Not moved Then
                If MsgBox("Create Project?", vbYesNo) = vbYes Then
                    newName = InputBox("New Project Name")
                    Set newFolder = oFolder.Folders.Add(newName)
                
                    olItem.UnRead = False
                          
                    olItem.Move newFolder
                
                    moved = True
                End If
                If Not moved Then
                    If MsgBox("Create Meeting?", vbYesNo) = vbYes Then
                        Set oFolder = GetFolderPath("\\" & USERNAME &"\Inbox\!System\Projects\!Meetings")
                        
                        newName = InputBox("New Meeting Name")
                        Set newFolder = oFolder.Folders.Add(newName)
                    
                        olItem.UnRead = False
                              
                        olItem.Move newFolder
                    
                        moved = True
                    End If
                End If
            End If
            
        End If

            
    Next
    
End Sub
