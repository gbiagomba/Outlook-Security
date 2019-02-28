'------------------------------------------------------
'Appaend the [encrypt]  email tag to the subject line
'------------------------------------------------------

Sub encrypt_Email()
    'Button for when email window is open
    
    Dim NewMail As MailItem, oInspector As Inspector
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
    
    Else
        Set NewMail = oInspector.CurrentItem
    End If

    Dim strTemp As String


    strTemp = "[encrypt] " & NewMail.Subject
        NewMail.Subject = strTemp
End Sub

Sub encrypt()
Set myFolder = Session.GetDefaultFolder(olFolderInbox)
Set myItem = myFolder.Items.Add("IPM.Note.mail")
myItem.Display
With myItem
.Subject = "[encrypt]"
.Display
End With
End Sub