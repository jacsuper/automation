
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