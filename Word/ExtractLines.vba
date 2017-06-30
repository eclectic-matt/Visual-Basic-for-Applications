'Picks up tagged lines and outputs at the bottom of the document
Public Sub extractItemsWith()
  m = InputBox("Please enter the identifying text: " & vbNewLine & "e.g. 'Action: '")
  If m = vbNo Or m = vbCancel Then
    Exit Sub
  Else
      strInput = m
      strOutput = extractItems(strInput)
      Call outputActions(strOutput)
  End If
End Sub
  
Public Function extractItems(strIdentifier)
  Dim strItems As String
  Dim intLength As Integer
  intLength = Len(strIdentifier)
  items = "Extracted '" & strIdentifier & "'(s):" & vbNewLine
  
  With ActiveDocument

      For Each singleLine In .Paragraphs
          lineText = singleLine.Range.Text
          ' Case insensitive
          checkLine = UCase(lineText)
          If Left(checkLine, intLength) = UCase(strIdentifier) Then
              items = items + Trim(lineText)
          End If
      Next

  End With
  
extractItems = items

End Function

Public Sub outputActions(strOutput)

 With ActiveDocument

      .Range.Collapse wdCollapseEnd
          .Range.InsertParagraphAfter
          .Range.InsertAfter strOutput
      .Range.Expand

  End With

End Sub
