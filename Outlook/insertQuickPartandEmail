
' This is a custom Quick Parts routine which takes
'   @ groupNum = an integer "group number" with a linked QuickPart
'   Quick Parts are named as "stt g1" to "stt g7"
' The routine inserts a QuickParts block into the end of your current email content
Sub insertSTTBuildingBlock(groupNum As Integer)

Dim oInspector As Inspector
Dim oDoc As Word.Document
Dim wordApp As Word.Application
Dim oTemplate As Word.Template
Dim oBuildingBlock As Word.BuildingBlock

Dim sttGrp As String
  sttGrp = "stt g" & groupNum

  Set oInspector = Application.ActiveInspector

  If oInspector.EditorType = olEditorWord Then

      Set oDoc = oInspector.WordEditor

      Set wordApp = oDoc.Application
      Set oTemplate = wordApp.Templates(1)
      Set oBuildingBlock = oTemplate.BuildingBlockEntries(sttGrp)

      wordApp.Selection.EndOf Unit:=wdStory, Extend:=wdMove

      oBuildingBlock.Insert wordApp.Selection.Range, True

  End If

End Sub

' This routine takes a generic confirmation email and reformats, inserts a quick part, and then displays the resulting email
Sub sttBookingAutoEmail()

Dim oMail As Outlook.MailItem
Dim strMessageBody As String
Dim intGroup As Integer

If ActiveExplorer.Selection.count > 1 Then
    m = MsgBox("More than one email selected." & vbNewLine & _
                "This tool currently processes one message at a time." & vbNewLine & _
                "Please try again with only one email highlighted." & vbNewLine & vbNewLine & _
                "Exiting the tool now!", vbCritical, "ERROR - Multiple Emails")
    Exit Sub
End If

With ActiveExplorer.Selection.Item(1)
    
    strMessageBody = .Body
    
    intStartE = InStr(1, .Body, "EMAIL: ") + 6
    intEndE = InStr(1, .Body, "WORK PHONE:")
        intLenE = intEndE - intStartE
        
    strStaffEmail = Mid(.Body, intStartE, intLenE)
    
    strNewSubject = Replace(.Subject, "New form submission: ", "")
    intGroup = CInt(Right(strNewSubject, 1))

End With

' Messy block removed for simplicity
' This basically formats based on the generic email content, not useful to others

closingComments = "<span style='font-family:""Calibri"",sans-serif'>SIGNATURE LINE</span>"

Set oMail = Application.CreateItem(olMailItem)

    oMail.BodyFormat = olFormatHTML
    
    oMail.Subject = strNewSubject
    oMail.To = strStaffEmail

    oMail.Body = strMessageBody
    
    oMail.Display
    
    ' INSERTS THE QUICK PART FOR THE NUMBERED STT GROUP!
    Call insertSTTBuildingBlock(intGroup)
    
    oMail.HTMLBody = oMail.HTMLBody & closingComments & vbNewLine & S

    oMail.Display
    Set oMail = Nothing

End Sub
