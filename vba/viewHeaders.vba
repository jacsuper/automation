
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