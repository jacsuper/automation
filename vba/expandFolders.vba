
'Loop through al subfolders and expand
Public Sub ExpandAllFolders()
  On Error Resume Next
  Dim Ns As Outlook.NameSpace
  Dim Folders As Outlook.Folders
  Dim CurrF As Outlook.MAPIFolder
  Dim F As Outlook.MAPIFolder
  Dim ExpandDefaultStoreOnly As Boolean

  ExpandDefaultStoreOnly = True

  Set Ns = Application.GetNamespace("Mapi")
  Set CurrF = Application.ActiveExplorer.CurrentFolder

  If ExpandDefaultStoreOnly = True Then
    Set F = Ns.GetDefaultFolder(olFolderInbox)
    Set F = F.Parent
    Set Folders = F.Folders
    LoopFolders Folders, True

  Else
    LoopFolders Ns.Folders, True
  End If

  DoEvents
  Set Application.ActiveExplorer.CurrentFolder = CurrF
End Sub

'Recursive Loop
Private Sub LoopFolders(Folders As Outlook.Folders, _
  ByVal bRecursive As Boolean _
)
  Dim F As Outlook.MAPIFolder

  For Each F In Folders
    Set Application.ActiveExplorer.CurrentFolder = F
    DoEvents

    If bRecursive Then
      If F.Folders.Count Then
        LoopFolders F.Folders, bRecursive
      End If
    End If
  Next
End Sub
