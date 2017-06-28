' A quick user-defined function to get a range input - used for many subsequent functions

Public Function getRangeInput(strPrompt as string) as string
  Set getRangeInput = Application.InputBox(strPrompt, default:=Selection.Address, type:=8)
End Function

' This subroutine fills down a range of cells to convert a "print-formatted" sheet into a "data-processing sheet".

' For example, if the print-formatted sheet was as follows:
'    1       2          3 
' A  Joe     Group 1    10:00 - 11:00
' B          Group 2    11:00 - 12:00
' C  Ann     Group 3    13:00 - 14:00
' D

' This subroutine would copy information down empty cells to convert A1:D3 into this:
'    1       2          3 
' A  Joe     Group 1    10:00 - 11:00
' B  Joe     Group 2    11:00 - 12:00
' C  Ann     Group 3    13:00 - 14:00
' D  Ann     Group 3    13:00 - 14:00

Sub fillDownRange()

Dim objFillRng As Range
Dim intStartRow As Integer
Dim intStartCol As Integer
Dim intNumRows As Integer
Dim intNumCols As Integer
Dim strTextToCopy As String
Dim intRowCount As Integer
Dim intColCount As Integer

' Get a range of cells to fill down
Set objFillRng = Application.InputBox("Select a range to copy down: ", Default:=Selection.Address, Type:=8)
' NOTE: made a separate function as this is used so commonly!
'Set objFillRng = getRangeInput("Please select a range to fill: ")

' Get the first row and column
intStartRow = objFillRng.Row
intStartCol = objFillRng.Column
' Get the number of rows and columns
intNumRows = objFillRng.Rows.count
intNumCols = objFillRng.Columns.count
' Select the first cell in the range
Cells(intStartRow, intStartCol).Select

'****** MAIN LOOP ***************
' Start the Column Counter from 1 to number
For intColCount = 0 To intNumCols - 1
    ' Start the ROW Counter from 1 to the number of rows - 1
    ' (as we are offsetting down a row each time)
    For intRowCount = 1 To intNumRows - 1
        ' Pick up the text from the active cell
        strTextToCopy = ActiveCell.Value
        'Debug.Print ActiveCell.Address & " --- " & strTextToCopy
        
        ' Offset down by one row
        ActiveCell.Offset(1, 0).Select
        ' If this cell is empty
        If IsEmpty(ActiveCell) Then
            ' Paste the text into this cell
            ActiveCell.Value = strTextToCopy
        Else
            ' Use this text to copy down
            strTextToCopy = ActiveCell.Value
        End If
    Next
    ' Select the next column along
    Cells(intStartRow, intStartRow + intColCount).Select
    'Debug.Print "Moving to " & ActiveCell.Address
Next

End Sub
