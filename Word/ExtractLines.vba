'---------------------------------------------------------------
'Picks up tagged lines and outputs at the bottom of the document
'---------------------------------------------------------------
' For the example document below
' If you call extractItemsWith and input "Action"
' You would get additional lines at the bottom with
  ' Action: FIRST
  ' Action: SECOND
  ' Action: Third
' Which helps Secretaries to process minutes quicker!
'---------------------------------------------------------------
' -- START OF DOCUMENT
'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam non neque ut dui pulvinar iaculis. Pellentesque porttitor nisl non nisi posuere ultricies. Aliquam metus metus, laoreet vitae diam eu, tincidunt fringilla dui. In eu sem urna. Vestibulum in turpis a odio rhoncus finibus. Aliquam venenatis massa sit amet efficitur laoreet. Donec ut ante neque. Donec pellentesque volutpat metus.
'Action: FIRST
'Praesent mattis varius ligula, et sagittis ipsum tincidunt vel. Suspendisse potenti. Aliquam iaculis gravida eros, ac placerat diam scelerisque a. Suspendisse tincidunt enim nulla, vel porttitor massa dignissim et. Vivamus in enim lectus. Praesent lobortis, mi at aliquet dapibus, sapien mi vestibulum neque, eu consectetur arcu magna non nisl. Nunc nunc ligula, elementum euismod justo eget, hendrerit porta lectus. Phasellus nisi leo, fermentum at blandit id, suscipit eget eros. Mauris mattis lacinia libero, nec tristique risus condimentum at. Etiam aliquam erat vel nisi luctus, a tristique dui porttitor. Aenean porttitor tincidunt elit, ultricies rutrum ligula mattis ut. Etiam eget ex vel tortor blandit sagittis. Vivamus consequat consequat consequat.
'Acting: as a test
'Fusce sed elit faucibus, dictum tellus sit amet, hendrerit erat. In vel commodo nulla, non suscipit quam. Nunc semper id risus sed dictum. Suspendisse ultrices convallis eleifend. Suspendisse aliquam cursus magna in egestas. Mauris aliquet, nisi nec aliquam mollis, mi velit gravida nisl, at tristique lacus diam a tortor. Maecenas ut vehicula diam, nec lobortis arcu. Fusce sed rhoncus risus. Aliquam eget mauris condimentum, sodales mi nec, auctor ligula. Nulla facilisi. Donec nec feugiat lorem.
'Action: SECOND
'Curabitur tortor augue, sagittis et magna at, luctus pharetra ligula. Etiam a elit massa. Suspendisse in nisi libero. Nam vitae lorem ac ex dignissim ultricies. Integer iaculis neque a velit egestas, ut pellentesque libero rutrum. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Aenean ultrices suscipit tortor a elementum. Maecenas consectetur, justo eu placerat sagittis, nisl enim commodo lacus, ut iaculis tortor lorem at diam. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Maecenas vel sollicitudin orci. In consectetur at libero quis varius. Vivamus tempor felis sed ullamcorper consequat.
'Phasellus elementum elementum bibendum. Nullam in mattis ex. Pellentesque tempor lacus ultricies, tincidunt odio id, facilisis nulla. Curabitur mollis accumsan metus, ac dignissim lorem mollis sed. Ut sollicitudin massa quis rutrum ultricies. Vestibulum varius felis sit amet suscipit imperdiet. Integer sit amet purus odio. Sed et faucibus nulla. Vestibulum vitae urna sem. Phasellus malesuada, quam sit amet sollicitudin vehicula, mauris massa vestibulum neque, lacinia auctor arcu tortor id nunc. Morbi facilisis posuere dui a suscipit. Nunc dolor nibh, sodales auctor placerat a, laoreet non eros. Aenean euismod quam sit amet turpis imperdiet dictum.
'ACTION: Third
' -- END OF DOCUMENT
'---------------------------------------------------------------

Public Sub extractItemsWith()
  m = InputBox("Please enter the identifying text: " & vbNewLine & "e.g. 'Action: '")
  If m = "" Then
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
