'----------------------------
' @name  showLastRow
' @descr Brings up a MsgBox showing the "last" row in column A of the current sheet
' @usage Run on a worksheet to show the number of filled rows in column A
'----------------------------
Sub showLastRow()
    lastRow = ActiveSheet.Cells(Rows.count, "A").End(xlUp).Row
    MsgBox ("Last Row: " & lastRow)
End Sub

'----------------------------
' @name  Col_Letter
' @arg   lngCol: Takes a column number (long)
' @descr Returns the column letter from A (1) - XFD (16384)
' @usage Use in VBA, or as =PERSONAL.XLSB!Col_Letter(COLUMN())
' @src   http://stackoverflow.com/a/12797190
'----------------------------
Function Col_Letter(lngCol As Long) As String
    Col_Letter = Split(Cells(, lngCol).Address, "$")(1)
End Function


' Usage as in "=PERSONAL.XLSB!findBlankCol(ROW(),TRUE)"
Function findBlankCol(Row As Integer, returnType As Boolean)
  Dim thisWb As Workbook
  Dim thisWs As Worksheet
  Dim BlankCol As Integer
  Dim BlankColName As String

  Set thisWb = ActiveWorkbook
  Set thisWs = thisWb.Sheets(ActiveSheet.Name)
  thisWs.Select

  With thisWs
      BlankCol = .Cells(Row, .Columns.count).End(xlToLeft).Column + 1
  End With

  Select Case (returnType)
      Case True
          'Number
          findBlankCol = BlankCol

      Case False
          'Text
          BlankColName = Chr(64 + BlankCol)
          findBlankCol = BlankColName
  End Select
End Function

Sub DeleteBlankRows()

  Dim Rng As Range
  Dim WorkRng As Range

  On Error Resume Next

  xTitleId = "Delete Blank Rows Tool"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)

  Application.ScreenUpdating = False

  For Each Rng In WorkRng

      If WorksheetFunction.CountA(Cells(Rng.Row, 1)) = 0 Then
          Rows(Rng.Row).Delete
      End If

  Next Rng

  Application.ScreenUpdating = True

  MsgBox ("END - Blank Rows Deleted")

End Sub

' Checks for a specific term "WORKSHOPS" and deletes all rows with this term in the first column
Sub DeleteWSRows()

  Dim Rng As Range
  Dim WorkRng As Range

  On Error Resume Next

  xTitleId = "Delete Workshops Tool"
  strToFind = "WORKSHOP"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)

  Application.ScreenUpdating = False

  For Each Rng In WorkRng

      If Cells(Rng.Row, 8).Text = strToFind Then
          Rows(Rng.Row).Delete
      End If

  Next Rng

  Application.ScreenUpdating = True

  MsgBox ("END - WS Deleted")

End Sub

' Copies down - mainly for testing
Sub CopyDown()
  Dim Rng As Range
  Dim WorkRng As Range
  Dim cellVal As String
  Dim Record As Integer
  'Dim total As Integer
  Dim PctDone As Single

  On Error Resume Next

  Complete = False
  tStart = Timer
  xTitleId = "Fill Down Cells Tool"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to fill down" & vbNewLine & "i.e. if a cell has a value, copy this to the cell below to allow filters to work correctly", xTitleId, WorkRng.Address, Type:=8)

  Application.ScreenUpdating = False
  'Application.ScreenUpdating = True
  Call ShowProgressForm

  Record = 0

  total = WorkRng.CountLarge
  'Call ProgBar(total)

  For Each Rng In WorkRng

      cellVal = Rng.Value
      If Len(cellVal) > 0 Then

         If Rng.Offset(RowOffset:=1, columnOffset:=0).Value = "" Then
         'Or Rng.Offset(rowOffset:=1, columnOffset:=0).Value Is Null
              Rng.Offset(RowOffset:=1, columnOffset:=0).Value = cellVal
              Record = Record + 2
              ' Update the percentage completed.
              PctDone = Record / total
              ' Call subroutine that updates the progress bar.
              Call UpdateProgressBar(PctDone, Record, total)
          End If

      End If

  Next Rng

  Rows("2:2").Select
  Selection.Copy
  Rows("3:500").Select
  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Columns("F:F").Select
  'Selection.NumberFormat = "0"
  Selection.NumberFormat = "############"

  Application.ScreenUpdating = True
  'Completed:
  t = Timer - tStart
  Complete = True
  With ProgressForm
      .FrameProgress.Caption = "COMPLETED."
      .LabelProgress.Caption = "TASK COMPLETE."
      .Label2.Caption = "SUCCESS - " & Denominator & " records updated."
      .Label1.Caption = "Time taken: approx. " & Round(t, 1) & " seconds."
  End With
  ProgressForm.Show
  t = Timer - tStart
  'MsgBox ("MACRO TOOK " & t & " secs")

  Call Unload(ProgressForm)

End Sub

' Highlights (colours) all blank cells (those without any text)
Sub HighlightBlanks()
  Dim Rng As Range
  Dim WorkRng As Range
  On Error Resume Next
  xTitleId = "Highlight Blanks Tool"
  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)
  Application.ScreenUpdating = False
  For Each Rng In WorkRng
      If Len(Rng.Text) < 1 Then
          Rng.Select
          With Selection.Interior
              .Pattern = xlSolid
              .PatternColorIndex = xlAutomatic
              .Color = 65535
              .TintAndShade = 0
              .PatternTintAndShade = 0
          End With
      End If
  Next Rng
  Application.ScreenUpdating = True
  MsgBox ("END - Highlighted")
End Sub

' Clears "blank" cells (those with no text)
Sub clearBlanks()
  Dim checkRng As Range
  Dim xTitleId As String
  Set checkRng = Application.Selection
  xTitleId = "Matt's Blank Clearer"
  Set checkRng = Application.InputBox("Select range to clear blanks:", xTitleId, checkRng.Address, Type:=8)
  Dim testCell As Object
  Dim testValue As String

  For Each testCell In checkRng

      With testCell

          .Select
          testValue = Selection.Value
          'MsgBox (testValue)
          If (testValue = "") Or (Selection = "") Or (testValue = 0) Then
              Selection.ClearContents
          End If

      End With

  Next

  MsgBox ("Cleared")
End Sub

' Another "fill down" function, I was having issues with XLS files from our data reporting portal!
Sub FillDownForFilters()
  Dim Rng As Range
  Dim Below As Range
  Dim WorkRng As Range
  Dim WS As Worksheet
  Dim off As Integer
  Dim cellVal As String

  debugging = True
  maxRowsToFill = 10

  On Error Resume Next

  xTitleId = "Fill Down Cells Tool"

  Set WorkRng = Application.Selection
  Set WorkRng = Application.InputBox("Select the range to fill down" & vbNewLine & "i.e. if a cell has a value, copy this to the cell below to allow filters to work correctly", xTitleId, WorkRng.Address, Type:=8)
  Set WS = Application.ActiveSheet

  'Application.ScreenUpdating = False

  For Each Rng In WorkRng

      If Len(Rng.Value) > 0 Then

          cellVal = Rng.Value
          If debugging Then
              Debug.Print "---------------------------"
              Debug.Print "Value to check = " & cellVal
          End If

          For off = 1 To maxRowsToFill

              Set Below = Rng.Offset(off, 0)
              belowVal = Below.Text

              If debugging Then
                  Debug.Print "Value below = " & belowVal
              End If

              If (belowVal = "") Or (belowVal Is Null) Then
                  Below = cellVal
                  If debugging Then
                      Debug.Print "Inserting " & cellVal
                  End If
              Else
                  GoTo nextOff
              End If

  nextOff:
          Next off

      End If

  Next Rng

  'Application.ScreenUpdating = True

  MsgBox ("END - Filled Down")

End Sub

' Testing - auto fills down a range
Sub AutoFillDown()
  Dim LR As Long
  LR = Range("A" & Rows.count).End(xlUp).Row
  Range("A" & LR).AutoFill Destination:=Range("A" & LR).Resize(2)
End Sub
