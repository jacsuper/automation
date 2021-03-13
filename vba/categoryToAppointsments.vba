Private WithEvents Items As Outlook.Items
Private WithEvents CalItems As Outlook.Items



Private Const USER = "surname"
Private Const USERNAME = "windowsfolderusername"


'Bind some events on startup
Private Sub Application_Startup()
    Dim folder As Outlook.folder
    
    'Bind new calendar entries
    Set CalItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    
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
