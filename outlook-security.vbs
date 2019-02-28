'-------------------------------------------------
'Forward an email as attachment
'-------------------------------------------------

Sub ForwardSpamToNetworkBox()

On Error Resume Next

Dim objItem As Outlook.MailItem

If Application.ActiveExplorer.Selection.Count = 0 Then
   MsgBox ("No item selected")
   Exit Sub
End If

For Each objItem In Application.ActiveExplorer.Selection
Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
        .Attachments.Add objItem, olEmbeddeditem
        .Subject = "Spear Phishing Attempt"
        .To = "your-it-security-email-goes-here@example.com"
        .Send
    End With
objItem.Delete
Next

Set objItem = Nothing
Set objMsg = Nothing

End Sub
