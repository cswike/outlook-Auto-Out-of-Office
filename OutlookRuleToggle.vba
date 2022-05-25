'Trigger with an appointment that sets off a reminder -
'must be named either "Enable Out of Office" or "Disable Out of Office"

Public WithEvents olRemind As Outlook.Reminders


Private Sub Application_Reminder(ByVal Item As Object)

    Set olRemind = Outlook.Reminders

    Dim i As Integer

    If Item.MessageClass <> "IPM.Appointment" Then Exit Sub

    If Item.Subject = "Enable Out of Office" Then
        Call OnOffRunRule("Out of Office", True, False)
        'Wait 5 seconds
        Wait (5)
        'Dismiss reminder
        Item.ReminderSet = False
        Item.Save

    ElseIf Item.Subject = "Disable Out of Office" Then
        Call OnOffRunRule("Out of Office", False, False)
        'Wait 5 seconds
        Wait (5)
        'Dismiss reminder
        Item.ReminderSet = False
        Item.Save
    End If

End Sub

'Enable or disable a rule
Sub OnOffRunRule(RuleName As String, Enable As Boolean, Optional blnExecute As Boolean = True)
    Dim olRules As Outlook.Rules
    Dim olRule As Outlook.Rule
    Dim intCount As Integer
 
    Set olRules = Application.Session.DefaultStore.GetRules
    Set olRule = olRules.Item(RuleName)
    
    If Enable Then olRule.Enabled = True Else olRule.Enabled = False
    
    If blnExecute Then olRule.Execute ShowProgress:=True
        olRules.Save
  
    Set olRules = Nothing
    Set olRule = Nothing
End Sub


'Delay seconds
Function Wait(nSeconds As Integer) As Boolean
    Dim dCurrentTime As Date
 
    dCurrentTime = Now
 
    Do Until DateAdd("s", nSeconds, dCurrentTime) <= Now
       DoEvents
    Loop
End Function
