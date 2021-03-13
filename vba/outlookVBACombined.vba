Private WithEvents Items As Outlook.Items
Private WithEvents CalItems As Outlook.Items
Private WithEvents MailItems As Outlook.Items



    Private Declare PtrSafe Function ShowWindow Lib "USER32" _
        (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Boolean
        
    Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetClassName Lib "USER32" Alias "GetClassNameA" _
                                     (ByVal hWnd As LongPtr, ByVal lpClassName As String, _
                                      ByVal nMaxCount As LongPtr) As LongPtr


Private Declare PtrSafe Function GetWindow Lib "USER32" _
                                  (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
                                  
Private Declare PtrSafe Function SetForegroundWindow Lib "USER32" (ByVal hWnd As Long) As LongPtr

Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As LongPtr

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"

'Find a folder by splitting a path
Function GetFolderPath(ByVal folderPath As String) As Outlook.folder
    Dim oFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
        
    On Error GoTo GetFolderPath_Error
    If Left(folderPath, 2) = "\\" Then
        folderPath = Right(folderPath, Len(folderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(folderPath, "\")
    Set oFolder = Application.Session.Folders.item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
        
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

'Bind some events on startup
Private Sub Application_Startup()
    Dim folder As Outlook.folder
    
    'Bind new calendar entries
    Set CalItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    
    'Bind new Inbox items
    Set MailItems = Session.GetDefaultFolder(olFolderInbox).Items
    
End Sub

'What are my categories again
Sub Explain()
    MsgBox "Red is for Today, Green is for Week, Blue is for Month"
End Sub

'Calendar item added
Private Sub CalItems_ItemAdd(ByVal item As Object)
    On Error Resume Next
    Dim Appt As Outlook.AppointmentItem
    If TypeOf item Is Outlook.AppointmentItem Then
        Set Appt = item
        If Appt.ReminderSet = False Then
            If MsgBox("NO REMINDER IS SET! Do you want to add one? " & item.Subject, vbYesNo) = vbYes Then
                Appt.ReminderSet = True
                Appt.ReminderMinutesBeforeStart = 15
                Appt.Save
            End If
        End If
        AddCategoryToPrivateAppointment Appt
        MakeSureOOTOSetAsFree Appt
    End If
End Sub

'Calendar item changed
Private Sub CalItems_ItemChange(ByVal item As Object)
    On Error Resume Next
    Dim Appt As Outlook.AppointmentItem
    If TypeOf item Is Outlook.AppointmentItem Then
        Set Appt = item

        AddCategoryToPrivateAppointment Appt
    End If
End Sub

'MakeSure OOTO Set As Free
Private Sub MakeSureOOTOSetAsFree(Appt As Outlook.AppointmentItem)
        If InStr(Appt.Subject, "PTO") <> 0 Or _
           InStr(Appt.Subject, "OOO") <> 0 Or _
           InStr(Appt.Subject, "OOTO") <> 0 Then
            ' Mark as Free
            If Not Appt.BusyStatus = olFree Then
                If MsgBox("OOTO with not Free.  Do you want to set as free?", vbYesNo) = vbYes Then
                    Appt.BusyStatus = olFree
                    Appt.Save
                End If
            End If
        End If
End Sub

'Categorize incoming meetings
Private Sub AddCategoryToPrivateAppointment(Appt As Outlook.AppointmentItem)
  ' #1 Private
  If Appt.Sensitivity = olPrivate Then
    If Len(Appt.Categories) = 0 Then
      Appt.Categories = "Private"
      Appt.Save
    End If
  ' #2 Private
  ElseIf InStr(1, Appt.Categories, "Private", vbTextCompare) Then
    Appt.Sensitivity = olPrivate
    Appt.Save
  ' #3 1:1
  ElseIf InStr(1, Appt.Subject, "1:1", vbTextCompare) Or InStr(1, Appt.Subject, "one on one", vbTextCompare) Then
    If Len(Appt.Categories) = 0 Then
      Appt.Categories = "Cal_1on1"
      Appt.Save
    End If
  ' #5 Travel
  ElseIf InStr(1, Appt.Subject, "travel", vbTextCompare) Or _
         InStr(1, Appt.Subject, "flight", vbTextCompare) Or _
         InStr(1, Appt.Subject, "hotel", vbTextCompare) Or _
         InStr(1, Appt.Subject, "trip", vbTextCompare) Or _
         InStr(1, Appt.Subject, "drive", vbTextCompare) Or _
         InStr(1, Appt.Subject, "commute", vbTextCompare) Or _
         InStr(1, Appt.Subject, "airport", vbTextCompare) Then
    If Len(Appt.Categories) = 0 Then
      Appt.Categories = "Cal_Travel"
      Appt.Save
    End If
  ' #4 From Me and not 1:1
  ElseIf InStr(1, Appt.Organizer, USER, vbTextCompare) Then
    If Len(Appt.Categories) = 0 Then
      Appt.Categories = "Cal_FromMe"
      Appt.Save
    End If
  End If
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

'String ends with string
Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

'String starts with string
Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

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

'Show mail headers
Sub ViewInternetHeader()
    Dim olItem As Outlook.MailItem, olMsg As Outlook.MailItem
    Dim strheader As String

    For Each olItem In Application.ActiveExplorer.Selection
        strheader = GetInetHeaders(olItem)
    
        Set olMsg = Application.CreateItem(olMailItem)
        With olMsg
            .BodyFormat = olFormatPlain
            .Body = strheader
            .Display
        End With
    Next
    Set olMsg = Nothing
End Sub

'Headers
Function GetInetHeaders(olkMsg As Outlook.MailItem) As String
    ' Purpose: Returns the internet headers of a message.'
    ' Written: 4/28/2009'
    ' Author:  BlueDevilFan'
    ' //techniclee.wordpress.com/
    ' Outlook: 2007'
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = olkMsg.PropertyAccessor
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Set olkPA = Nothing
End Function

'Clear Deleted
Sub RemoveAllItemsAndFoldersInDeletedItems()
    Dim oDeletedItems As Outlook.folder
    Dim oFolders As Outlook.Folders
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    Set oItems = oDeletedItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.item(i).Delete
    Next
    Set oFolders = oDeletedItems.Folders
    For i = oFolders.Count To 1 Step -1
        oFolders.item(i).Delete
    Next
End Sub

'Open new mail to self, clear signatures, set focus to subject
Sub NewMailToSelf()
    Dim objMsg As MailItem
    Dim item As Object

    Set objMsg = Application.CreateItem(olMailItem)
    
    With objMsg
      .To = "me@email.com"
      .CC = ""
      .BCC = ""
      .Subject = " Notes"
      .BodyFormat = olFormatHTML
      .Body = ""
        
      .Display
    
    
    End With
    
    DoEvents
    
    If SetFocusOnBody(Application.ActiveInspector.Caption) = True Then
        ' focus is set
    End If
    
    SendKeys "+{TAB}"
    
    Set objMsg = Nothing

End Sub

'Generate quick follow up mail
Public Sub FollowUp()
    Dim objMsg As MailItem
    Dim Selection As Selection
    Dim obj As Object
    
    Set Selection = ActiveExplorer.Selection
    
    For Each obj In Selection
    
    Set objMsg = Application.CreateItem(olMailItem)
    
     With objMsg
      .To = obj.SenderName
      .Subject = ""
      
      For Each objOutlookRecip In objMsg.Recipients
        objOutlookRecip.Resolve
      Next
      
      .Display
    
    End With
    Set objMsg = Nothing
    
    Next

End Sub



'-----------------------------------------------------------------------
' Module :modBodyFocus
' DateTime :02.02.2004 09:01
' Author :Michael Bauer
' Purpose :Sets focus on an OL Inspector body.
' Tested :OL2k
'-----------------------------------------------------------------------
Private Function FindChildClassName(ByVal lHwnd As LongPtr, _
sFindName As String _
) As LongPtr
Dim lRes As LongPtr

lRes = GetWindow(lHwnd, GW_CHILD)
If lRes Then
Do
If GetClassNameEx(lRes) = sFindName Then
FindChildClassName = lRes
Exit Function
End If
lRes = GetWindow(lRes, GW_HWNDNEXT)
Loop While lRes <> 0
End If
End Function

Public Function GetBodyHandle(ByVal lInspectorHwnd As Long) As LongPtr
Dim lRes As LongPtr

lRes = FindChildClassName(lInspectorHwnd, "AfxWnd")
If lRes Then
lRes = GetWindow(lRes, GW_CHILD) ' OL2000: ClassName = "#32770"
If lRes Then
lRes = FindChildClassName(lRes, "AfxWnd")
If lRes Then
lRes = GetWindow(lRes, GW_CHILD)
If lRes Then
' plain/text: ClassName="RichEdit20A", html:
ClassName = "Internet Explorer_Server"
GetBodyHandle = GetWindow(lRes, GW_CHILD)
End If
End If
End If
End If
End Function

Private Function GetClassNameEx(ByVal lHwnd As LongPtr) As String
Dim lRes
Dim sBuffer As String * 256
lRes = GetClassName(lHwnd, sBuffer, 256)
If lRes <> 0 Then
GetClassNameEx = Left$(sBuffer, lRes)
End If
End Function

Public Function GetInspectorHandle(ByVal sCaption As String) As LongPtr
GetInspectorHandle = FindWindow("rctrl_renwnd32", sCaption)
End Function

Public Function SetFocusOnBody(sInspectorCaption As String) As Boolean
Dim lHwnd

lHwnd = GetInspectorHandle(sInspectorCaption)
lHwnd = GetBodyHandle(lHwnd)
If lHwnd Then
SetFocusOnBody = CBool(SetForegroundWindow(lHwnd))
End If
End Function

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

'Move all in folder to Deleted Items
Private Function DeleteFolderContents(folderPath As String) As Boolean
    Dim oDeletedItems As Outlook.folder
    Dim oFolder As Outlook.folder
    Dim oItems As Outlook.Items
    Dim i As Long
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    Set oFolder = GetFolderPath(folderPath)
    
    If Not oFolder Is Nothing Then
        Set oItems = oFolder.Items
        For i = oItems.Count To 1 Step -1
            oItems.item(i).Move oDeletedItems
        Next
        DeleteFolderContents = True
        Exit Function
    End If
    
    DeleteFolderContents = False

  
End Function
'Mark all in folder as read
Private Function MarkAsReadFolderContents(folderPath As String) As Boolean
    Dim oFolder As Outlook.folder
    Dim oItems As Outlook.Items
    Dim i As Long

    Set oFolder = GetFolderPath(folderPath)
    
    If Not oFolder Is Nothing Then
        Set oItems = oFolder.Items
        For i = oItems.Count To 1 Step -1
            If (oItems.item(i).UnRead = True) Then
                oItems.item(i).UnRead = False
            End If
        Next
        MarkAsReadFolderContents = True
        Exit Function
    End If
    
    MarkAsReadFolderContents = False

  
End Function

'Move all in my junk folders to Deleted
Public Sub ClearJunkFolders()
    DeleteFolderContents ("\\" & USERNAME & "\Inbox\!System\KB\KPI")

End Sub

'Macro to move to deleted items after setting category to nothing
Public Sub DeleteSelected()
    Dim oDeletedItems As Outlook.folder
    
    Dim olItem As Outlook.MailItem
    
    'Obtain a reference to deleted items folder
    Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Do you REALLY want to delete all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
        If Response = vbNo Then Exit Sub
    End If
    
    For Each olItem In Application.ActiveExplorer.Selection

        olItem.UnRead = False
        olItem.Categories = ""
        olItem.Save
        
        olItem.Move oDeletedItems
 
    Next
  

End Sub

'Clear categories and flag
Public Sub ClearCategoriesAndFlag()
    Dim olItem As Outlook.MailItem
    
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Do you REALLY want to categorize all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
        If Response = vbNo Then Exit Sub
    End If
    
    For Each olItem In Application.ActiveExplorer.Selection

        olItem.UnRead = False
        olItem.Categories = ""
        olItem.Save
        
        If Not olItem.IsMarkedAsTask Then
            olItem.MarkAsTask olMarkComplete
        End If
            
        olItem.Save
            
        olItem.FlagStatus = olFlagComplete
        
        olItem.Save
 
    Next
End Sub

'Do Done/Archive
Private Function DoDone(olItem As Object, oArchive As Outlook.folder) As Boolean

    olItem.UnRead = False
    olItem.Categories = "Done"
                
    If TypeOf olItem Is Outlook.MailItem Then
        If Not olItem.IsMarkedAsTask Then
            olItem.MarkAsTask olMarkComplete
        End If
    End If
        
    olItem.Save
        
    olItem.FlagStatus = olFlagComplete
        
    olItem.Save
        
    olItem.Move oArchive
 

    DoDone = True
End Function

'Active selected - be it in a list of open email
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
       
    Set objApp = Nothing
End Function

'Active Window
Function GetActiveSelectionType() As String
    Dim objApp As Outlook.Application
           
    GetActiveSelectionType = ""
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            GetActiveSelectionType = "List"
        Case "Inspector"
            GetActiveSelectionType = "Email"
    End Select
       
    
End Function

'Archive email
Public Sub SimulateDoneQuickStep()
    Dim oArchive As Outlook.folder
    Dim oInspector As Inspector
    
    Dim olItem As Object
    
    'Obtain a reference to archive items folder
    Set oArchive = Session.GetDefaultFolder(olFolderInbox).Folders("!System").Folders("Archive")
    
    Dim windowType As String
    
    windowType = GetActiveSelectionType()
    
    If windowType = "List" Then
        If Application.ActiveExplorer.Selection.Count > 1 Then
            Dim Response As VbMsgBoxResult
            Response = MsgBox("Do you REALLY want to done all SELECTED ?" _
                            , vbExclamation + vbDefaultButton2 + vbYesNo)
            If Response = vbNo Then Exit Sub
        End If
        
        For Each olItem In Application.ActiveExplorer.Selection
          DoDone olItem, oArchive
        Next
        
    End If
    
    If windowType = "Email" Then
        Set oInspector = Application.ActiveInspector
        
        If Not oInspector Is Nothing Then
            Set olItem = oInspector.CurrentItem
            
            If Not olItem Is Nothing Then
                DoDone olItem, oArchive
            End If
        End If
        
    End If

End Sub

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

'OOTO
Public Sub OOTO()
    Dim objAppt As AppointmentItem

    Set objAppt = Application.CreateItem(olAppointmentItem)

    With objAppt
        .RequiredAttendees = "email@email.com"
        .Subject = "Me (me@) OOTO"
        .Location = ""
        .Body = ""
        .MeetingStatus = 1
        .ResponseRequested = False
        .BusyStatus = olFree
        .Recipients.ResolveAll
        .AllDayEvent = True
        .ReminderSet = False
        
        .Display
    
    End With

    
End Sub

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

'Get login
Private Function StripLogin(email As String) As String
    StripLogin = Left(email, InStr(1, email, "@") - 1)
End Function

'Helper to resolve
Function GetSMTPAddress(ByVal strAddress As String)
' Based on:
' http://blogs.msdn.com/vikas/archive/2007/10/24/oom-getting-primary-smtp-address-from-x400-x500-sip-ccmail-etc.aspx
Dim olApp As Object
Dim oCon As Object
Dim strKey As String
Dim oRec As Object
Dim strRet As String
Dim fldr As Object
    
    If InStr(strAddress, "@") > 0 Then
        GetSMTPAddress = strAddress
    Else
        'IF OUTLOOK VERSION IS >= 2007 THEN USES NATIVE OOM PROPERTIES AND METHODS
        On Error Resume Next
        Set olApp = Application
        Set fldr = olApp.GetNamespace("MAPI").GetDefaultFolder(10).Folders.item("Random")
        If fldr Is Nothing Then
            olApp.GetNamespace("MAPI").GetDefaultFolder(10).Folders.Add "Random"
            Set fldr = olApp.GetNamespace("MAPI").GetDefaultFolder(10).Folders.item("Random")
        End If
        On Error GoTo 0
        If CInt(Left(olApp.Version, 2)) >= 12 Then
            Set oRec = olApp.Session.CreateRecipient(strAddress)
            If oRec.Resolve Then
                strRet = oRec.AddressEntry.GetExchangeUser.PrimarySmtpAddress
            End If
        End If
        If Not strRet = "" Then GoTo ReturnValue
        'IF OUTLOOK VERSION IS < 2007 THEN USES LITTLE HACK
        'How it works
        '============
        '1) It will create a new contact item
        '2) Set it's email address to the value passed by you, it could be X500,X400 or any type of email address stored in the AD
        '3) We will assign a random key to this contact item and save it in its Fullname to search it later
        '4) Next we will save it to local contacts folder
        '5) Outlook will try to resolve the email address & make AD call if required else take the Primary SMTP address from its cache and append it to Display name
        '6) The display name will be something like this " ( email.address@server.com )"
        '7) Now we need to parse the Display name and delete the contact from contacts folder
        '8) Once the contact is deleted it will go to Deleted Items folder, after searching the contact using the unique random key generated in step 3
        '9) We then need to delete it from Deleted Items folder as well, to clean all the traces
        Set oCon = fldr.Items.Add(2)
        oCon.Email1Address = strAddress
        strKey = "_" & Replace(Rnd * 100000 & Format(Now, "DDMMYYYYHmmss"), ".", "")
        oCon.FullName = strKey
        oCon.Save
        strRet = Trim(Replace(Replace(Replace(oCon.Email1DisplayName, "(", ""), ")", ""), strKey, ""))
        oCon.Delete
        Set oCon = Nothing
        Set oCon = olApp.Session.GetDefaultFolder(3).Items.Find("[Subject]=" & strKey)
        If Not oCon Is Nothing Then oCon.Delete
ReturnValue:
        GetSMTPAddress = strRet
    End If
    
End Function

'Resolve contact to smtp
Private Function EmailFromContact(contact As Outlook.ContactItem) As String
    Dim exchangeUser As Outlook.exchangeUser
    EmailFromContact = "user"
    
    If contact.Email1AddressType = "SMTP" Then
        EmailFromContact = contact.Email1Address
    Else
        EmailFromContact = GetSMTPAddress(contact.Email1Address)
    End If
End Function

'Resolve exchange user to smtp
Private Function EmailFromRecipient(recipient As Outlook.recipient) As String
    Dim exchangeUser As Outlook.exchangeUser
    EmailFromRecipient = "user"
    
    Set exchangeUser = recipient.AddressEntry.GetExchangeUser()
        If Not exchangeUser Is Nothing Then
             EmailFromRecipient = exchangeUser.PrimarySmtpAddress
        Else
             EmailFromRecipient = recipient.address
        End If
End Function

'Load external styles and intern
Private Function InlineStyles(html As String) As String
    Dim htm As Object
    Set htm = CreateObject("htmlfile")
    Dim style As Object
    Dim url As String
    Dim styleText As String
    Dim fullStyles As String
    
    
    fullStyles = ""
    InlineStyles = html

    htm.Write html

    
        For Each style In htm.getElementsByTagName("link")
            url = style.href
            
             With CreateObject("msxml2.xmlhttp")
                .Open "GET", url, False
                .Send
                styleText = .responseText
             End With
            
            fullStyles = fullStyles & styleText
        Next style
    InlineStyles = Replace(InlineStyles, "</head>", "<style>" & fullStyles & "</style></head>")
End Function
' this function queries a remote server using the structured web query it was passed
Public Function GetWebQuery(ByRef oXML, ByVal address)
 ' variables used in this function:
 Dim serverTime
  ' error processing
    GetWebQuery = False
  On Error Resume Next ' We only do this because of the inherent instability of queries
    Err.Clear
    ' create a GET request to the web site�
  oXML.Open "GET", address, False
    ' set whatever headers you want or need�
  oXML.setRequestHeader "Accept-Language", "en"
  oXML.setRequestHeader "Accept", "text/html"
  oXML.setRequestHeader "Accept", "application/json"
  ' send the request
  oXML.Send
    ' flag the result as true only if no errors occurred
    If Err.Number = 0 Then GetWebQuery = True
    ' automatically give the request a second chance in case it failed the first time
  If oXML.Status >= 400 And oXML.Status <= 599 Then
      oXML.Open "GET", address, False
      oXML.setRequestHeader "Accept-Language", "en"
      oXML.setRequestHeader "Accept", "text/html"
      oXML.Send
      ' flag the result as true only if no errors occurred
      If Err.Number = 0 Then GetWebQuery = True
    End If
 ' We can deal with response header, too �
 serverTime = oXML.getResponseHeader("Date")
    ' normal operation
    On Error GoTo 0
End Function

Sub OpenAppointmentCopy()
'=================================================================
'Description: Outlook macro to create a new appointment with
'             specific details of the currently selected
'             appointment and show it in a new window.
'
' author : Robert Sparnaaij
' version: 1.0
' website: https://www.howto-outlook.com/howto/openapptcopy.htm
'
' updated: 2016-07-02
' by     : dmmartin@
'=================================================================

    Dim objOL As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Dim personalAttachmentFilePath As String
    Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.recipient
    Set objOL = Outlook.Application
    personalAttachmentFilePath = ""
    
    'Get the selected item
    Select Case TypeName(objOL.ActiveWindow)
        Case "Explorer"
            Set objSelection = objOL.ActiveExplorer.Selection
            If objSelection.Count > 0 Then
                Set objItem = objSelection.item(1)
            Else
                Result = MsgBox("No item selected. " & _
                            "Please make a selection first.", _
                            vbCritical, "OpenAppointmentCopy")
                Exit Sub
            End If
        
        Case "Inspector"
            Set objItem = objOL.ActiveInspector.CurrentItem
            
        Case Else
            Result = MsgBox("Unsupported Window type." & _
                        vbNewLine & "Please make a selection" & _
                        "in the Calendar or open an item first.", _
                        vbCritical, "OpenAppointmentCopy")
            Exit Sub
    End Select

    Dim olAppt As Outlook.AppointmentItem
    Dim olApptCopy As Outlook.AppointmentItem
    Set olApptCopy = Outlook.CreateItem(olAppointmentItem)
    Dim myAttachments As Outlook.Attachments
        
        
    'Copy the desired details to a new appointment item
    If objItem.Class = olAppointment Then
        Set olAppt = objItem
        Set myAttachments = olAppt.Attachments
        
        
        With olApptCopy
            .MeetingStatus = olMeeting
            .Subject = olAppt.Subject
            .Location = olAppt.Location
            .Body = olAppt.Body
            .Categories = olAppt.Categories
            .AllDayEvent = olAppt.AllDayEvent
            .ReminderMinutesBeforeStart = olAppt.ReminderMinutesBeforeStart
            .RequiredAttendees = olAppt.RequiredAttendees
            .OptionalAttendees = olAppt.OptionalAttendees
            
            ' For Each olAttch In myAttachments
            '    Att = olAttch.FileName
            '    .Attachments.Add (Att)
            ' Next olAttch
            
            .Duration = olAppt.Duration
        End With
        
        'Display the copy
        olApptCopy.Display
    
    'Selected item isn't an appointment item
    Else
        Result = MsgBox("No appointment item selected. " & _
                    "Please make a selection first.", _
                    vbCritical, "OpenAppointmentCopy")
        Exit Sub
    End If
    
    'Clean up
    Set objOL = Nothing
    Set objItem = Nothing
    Set olAppt = Nothing
    Set olApptCopy = Nothing
    
End Sub

'Handle the default template parameters
Function FetchDefault(default As String, message As Object) As String
    Dim objOutlookRecip
    FetchDefault = ""
    If (InStr(default, ",") > 0) Then
        FetchDefault = Split(default, ",")(1)
        
        message.Recipients.ResolveAll
        If (FetchDefault = "TO") Then
            If (Not message.To = "") Then
            
              For Each objOutlookRecip In message.Recipients
                If (objOutlookRecip.Type = olTo) Then
                    FetchDefault = Trim(Split(objOutlookRecip.name, ",")(1))
                End If
              Next
            End If
        Else
            If (FetchDefault = "CC") Then
                If (Not message.CC = "") Then
                    For Each objOutlookRecip In message.Recipients
                    If (objOutlookRecip.Type = olCC) Then
                         FetchDefault = Trim(Split(objOutlookRecip.name, ",")(1))
                    End If
              Next
                End If
            End If
        End If
    End If
End Function

'Fill in a templat item
Public Sub FillInTemplate()
    Dim mail As Object
    
    Dim M1 As MatchCollection
    Dim M As Object
    Dim EAI As String
    Dim ishtml As Boolean
    
    Set mail = GetCurrentItem()
    
    Dim reg1 As Object
    Dim def
    Set reg1 = New RegExp

    With reg1
        .Pattern = "({)({)([a-zA-Z _0-9,.<>:\/=""]+)(})(})" ''All between {{}}
        .Global = False
        
        
    End With
        
       ' mail.Body = mail.HTMLBody
    If Not mail Is Nothing Then
        If TypeOf mail Is Outlook.MailItem Or TypeOf mail Is Outlook.AppointmentItem Then
            'Parse template items
            'Get tempate values
            'fill template
            Do While reg1.Test(mail.Subject)
                Set M1 = reg1.Execute(mail.Subject)
                For Each M In M1
                    EAI = M.SubMatches(2)
                    def = FetchDefault(EAI, mail)
                    newValue = InputBox(EAI, "Fill Template", def)
                    mail.Subject = Replace(mail.Subject, "{{" & EAI & "}}", newValue)
                Next
            Loop
            
            ishtml = False
            If TypeOf mail Is Outlook.MailItem Then
                ishtml = mail.BodyFormat = olFormatHTML
            End If
            
            If TypeOf mail Is Outlook.AppointmentItem Or Not ishtml Then
                'Parse template items
                'Get tempate values
                'fill template
                Do While reg1.Test(mail.Body)
                    Set M1 = reg1.Execute(mail.Body)
                    For Each M In M1
                        EAI = M.SubMatches(2)
                        def = FetchDefault(EAI, mail)
                        newValue = InputBox(EAI, "Fill Template", def)
                        mail.Body = Replace(mail.Body, "{{" & EAI & "}}", newValue)

                    Next
                Loop
            Else
                'Parse template items
                'Get tempate values
                'fill template
                Do While reg1.Test(mail.HTMLBody)
                    Set M1 = reg1.Execute(mail.HTMLBody)
                    For Each M In M1
                        EAI = M.SubMatches(2)
                        def = FetchDefault(EAI, mail)
                        newValue = InputBox(EAI, "Fill Template", def)
                        mail.HTMLBody = Replace(mail.HTMLBody, "{{" & EAI & "}}", newValue)
                    Next
                Loop
            End If
            
            mail.Save
            
        End If
    End If

End Sub

'Send mail in 5 mins
Public Sub DelaySend()
    Dim olItem As Outlook.MailItem
    
    
    If Application.ActiveInspector Is Nothing Then
        Exit Sub
    End If
    
    Set olItem = Application.ActiveInspector.CurrentItem
    
    If Not olItem.Sent Then
        olItem.DeferredDeliveryTime = DateAdd("n", 5, Now)
        olItem.Send
    End If
End Sub

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


