


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
