
'Quick nav to Inbox
Sub OpenInbox()

 Dim objOlApp As Outlook.Application
 Dim objFolder As Outlook.folder
 Set objOlApp = CreateObject("Outlook.Application")

 Set objFolder = Session.GetDefaultFolder(olFolderInbox)

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 Set objFolder = Nothing
 Set objOlApp = Nothing

End Sub

'Quick nav to Actions
Sub OpenActions()

 Dim objOlApp As Outlook.Application
 Dim objFolder As Outlook.folder
 Set objOlApp = CreateObject("Outlook.Application")

 Set objFolder = Session.GetDefaultFolder(olFolderInbox).Folders("!System").Folders("Actions")

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 Set objFolder = Nothing
 Set objOlApp = Nothing

End Sub

'Quick nav to Projects
Sub OpenProjects()

 Dim objOlApp As Outlook.Application
 Dim objFolder As Outlook.folder
 Set objOlApp = CreateObject("Outlook.Application")

 Set objFolder = Session.GetDefaultFolder(olFolderInbox).Folders("!System").Folders("Projects")

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 DoEvents

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder.Folders.GetFirst
 
 DoEvents
  
 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 Set objFolder = Nothing
 Set objOlApp = Nothing

End Sub

'Quick nav to Meetings
Sub OpenMeetings()

 Dim objOlApp As Outlook.Application
 Dim objFolder As Outlook.folder
 Set objOlApp = CreateObject("Outlook.Application")

 Set objFolder = Session.GetDefaultFolder(olFolderInbox).Folders("!System").Folders("Projects").Folders("!Meetings")

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 DoEvents

 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder.Folders.GetFirst
 
 DoEvents
  
 Set objOlApp.ActiveExplorer.CurrentFolder = objFolder

 
 Set objFolder = Nothing
 Set objOlApp = Nothing

End Sub

'Jump to a partial match folder quick
Public Sub JumpFolder()
    On Error Resume Next
 
    Dim folderToJump As String

    folderToJump = InputBox("Folder?")
    
    folderToJump = StripFolderNameOrSubject(folderToJump)
        
    If (folderToJump = "") Then
        Exit Sub
    End If
        
    Dim Ns As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim CurrF As Outlook.MAPIFolder
    Dim F As Outlook.MAPIFolder
    Dim ExpandDefaultStoreOnly As Boolean
    Dim bDidChange As Boolean
    
    bDidChange = False

    ExpandDefaultStoreOnly = True

    Set Ns = Application.GetNamespace("Mapi")
    Set CurrF = Application.ActiveExplorer.CurrentFolder

    If ExpandDefaultStoreOnly = True Then
        Set F = Ns.GetDefaultFolder(olFolderInbox)
        Set F = F.Parent
        Set Folders = F.Folders
        LoopFoldersJump Folders, True, False, bDidChange, folderToJump
    Else
        LoopFoldersJump Ns.Folders, True, False, bDidChange, folderToJump
    End If
  
    If Not bDidChange Then
        If ExpandDefaultStoreOnly = True Then
            Set F = Ns.GetDefaultFolder(olFolderInbox)
            Set F = F.Parent
            Set Folders = F.Folders
            LoopFoldersJump Folders, True, True, bDidChange, folderToJump
        Else
            LoopFoldersJump Ns.Folders, True, True, bDidChange, folderToJump
        End If
    End If
    
    

End Sub

'Recursive Loop for Jump
Private Sub LoopFoldersJump(Folders As Outlook.Folders, ByVal bRecursive As Boolean, ByVal bPartial As Boolean, ByRef bDidChange As Boolean, folderToJump As String)
  Dim F As Outlook.MAPIFolder
  Dim strFolderName As String

  If Not bDidChange Then
      For Each F In Folders
        strFolderName = StripFolderNameOrSubject(F.name)
        
        If Not bPartial Then
            If IsFolderAndSubjectEquivalent(folderToJump, strFolderName) Then
                Set Application.ActiveExplorer.CurrentFolder = F
                bDidChange = True
                Exit Sub
            End If
        Else
            If IsFolderAndSubjectEquivalent(folderToJump, strFolderName) Or IsFolderAndSubjectRelated(folderToJump, strFolderName) Then
                Set Application.ActiveExplorer.CurrentFolder = F
                bDidChange = True
                Exit Sub
            End If
        End If
    
        If bRecursive Then
          If F.Folders.Count Then
            LoopFoldersJump F.Folders, bRecursive, bPartial, bDidChange, folderToJump
          End If
        End If
      Next
  End If
End Sub