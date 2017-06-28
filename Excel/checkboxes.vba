' A bunch of functions to create and manage checkboxes in your Excel files!

Sub InsertCheckBoxes()
  Dim Rng As Range
  Dim WorkRng As Range
  Dim WS As Worksheet
  ' Not used currently:
  Dim chk As CheckBox

  On Error Resume Next

  xTitleId = "Matt's Checkbox Tool"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to add checkboxes to:", xTitleId, WorkRng.Address, Type:=8)
  Set WS = Application.ActiveSheet

  Application.ScreenUpdating = False

  For Each Rng In WorkRng
      With WS.CheckBoxes.Add(Rng.Left, Rng.Top, Rng.Width, Rng.Height)
          .Characters.Text = Rng.Value
      End With
  Next

  ' To link checkboxes (unclear if necessary for this purpose)
  For Each chk In ActiveSheet.CheckBoxes
     With chk
        .LinkedCell = .TopLeftCell.Offset(0, 0).Address
     End With
  Next chk

  WorkRng.ClearContents
  WorkRng.Select
  Application.ScreenUpdating = True

  MsgBox ("Checkboxes Inserted")
End Sub


Sub CheckBoxCull()
  Dim chk As CheckBox
  Dim i As Integer
  i = 0
  For Each chk In ActiveSheet.CheckBoxes
        chk.Delete
        i = i + 1
  Next chk
  MsgBox ("Checkbox Cull Complete. " & i & " boxes deleted.")
End Sub

Sub UncheckAll()
  Dim chk As CheckBox
  Dim i As Integer
  i = 0
  For Each chk In ActiveSheet.CheckBoxes
        chk.Value = False
        i = i + 1
  Next chk
  MsgBox ("UNChecking Complete. " & i & " boxes UNticked.")
End Sub

Sub CheckToggleRange()
  Dim isect As Object
  Dim Rng As Range
  Dim WorkRng As Range
  Dim WS As Worksheet
  Dim chk As CheckBox

  On Error Resume Next

  xTitleId = "Matt's Check Toggler"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the toggle checkboxes:", xTitleId, WorkRng.Address, Type:=8)
  Set WS = Application.ActiveSheet

  Application.ScreenUpdating = False

  For Each chk In ActiveSheet.CheckBoxes

      Set isect = Application.Intersect(WorkRng, Range(chk.LinkedCell))
      If isect Is Nothing Then
          'MsgBox ("Ranges do not intersect at " & chk.LinkedCell)
          GoTo skip:
      End If

  chk.Delete

  skip:
  Set isect = Nothing
  Next chk

  MsgBox ("Toggled checkboxes")
End Sub

Sub CheckAll()
  Dim chk As CheckBox
  Dim i As Integer
  i = 0
    For Each chk In ActiveSheet.CheckBoxes
          'If Left(chk.LinkedCell, 2) = "$A" Then
            chk.Value = True
            i = i + 1
          'End If
    Next chk
    MsgBox ("Checking Complete. " & i & " boxes ticked.")
End Sub

Sub ShowChecked()
  Dim m As Variant
  Dim Name As String
  Dim Invitees(100) As String
  Dim Idx As Integer
  Dim total As Integer
  Idx = 0
  Name = "" & vbNewLine & ""

  m = MsgBox("This will display each checked box in a list (this might take a while)." & vbNewLine & vbNewLine & "Continue at your own peril!!!", vbOKCancel)
  returnCol = 7

  If m = vbCancel Then
      Exit Sub
  End If

  For Each chk In ActiveSheet.CheckBoxes
      With chk
          If .Value = 1 Then
              Invitees(Idx) = Cells(Range(.LinkedCell).Row, returnCol).Value
              Idx = Idx + 1
          End If
      End With
  Next
  total = Idx
  ' List all invitees in MsgBox
  For Idx = 0 To total
      Name = Name & vbNewLine & Invitees(Idx)
  Next
  MsgBox ("Send invites to " & Name)
End Sub

