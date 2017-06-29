' Get the email address of the current user
Private Function getMyEmail()

    Dim olNS As Outlook.NameSpace
    Dim olFol As Outlook.Folder
    Dim myEmail As String
    
    Set olNS = Outlook.GetNamespace("MAPI")
    Set olFol = olNS.GetDefaultFolder(olFolderInbox)
    myEmail = olNS.Accounts.Item(1).SmtpAddress

    getMyEmail = myEmail
    
End Function


' Take the selected item and send in all formats to yourself to test - TO BE UPDATED!
Private Sub sendMailAllFormats()

myEmail = getMyEmail
MsgBox myEmail

'With Application.ActiveExplorer

End Sub
