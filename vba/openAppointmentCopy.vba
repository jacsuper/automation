
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