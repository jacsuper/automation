

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