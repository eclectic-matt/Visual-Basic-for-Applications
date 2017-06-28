'----------------------------
' @name  CloseAllWorkbooks
' @descr Closes ALL open workbooks without saving
' @usage No parameters, just closes everything without saving
'----------------------------
Sub CloseAllWorkbooks()
    For Each wb In Application.Workbooks
        wb.Close False
    Next
End Sub


'----------------------------
' @name  openNamedSheet
' @descr Opens the worksheet (in the active workbook) with the user entered sheet name
' @usage Run in Workbook and enter a sheet name in the InputBox
'----------------------------
Sub openNamedSheet()
    m = InputBox("Please enter a sheet name", "Open Named Sheet")
    If Not m = vbNo Or m = vbCancel Or Len(m) = 0 Then
        ActiveWorkbook.Sheets(m).Activate
    Else
        MsgBox "Cancelled"
    End If
End Sub


'----------------------------
' @name  complexAlphabetise
' @descr Sorts worksheets alphabetically by testing each one against each other
' @usage Run on a worksheet to sort - but not very efficient! Takes O(n)^2 - 1 to complete!!!!
' @src   https://www.extendoffice.com/documents/excel/629-excel-sort-sheets.html
'----------------------------
Sub complexAlphabetise()
    For i = 1 To moduleWb.Sheets.count
        For j = 1 To moduleWb.Sheets.count - 1
            If UCase$(moduleWb.Sheets(j).Name) > UCase$(moduleWb.Sheets(j + 1).Name) Then
                moduleWb.Sheets(j).Move After:=moduleWb.Sheets(j + 1)
            End If
        Next
    Next
End Sub

'----------------------------
' @name  sortSheets
' @descr Sorts worksheets alphabetically by listing sheet names in a new WS then using a sort function
' @usage Run on a worksheet to sort - more efficient for large # of worksheets!
' @src   http://www.excelforum.com/l/359252-sort-excel-sheets-into-alphabetical-order.html
'----------------------------
Sub SortSheets()
    Call startTimer
    Dim sht As Worksheet
    Dim mySht As Worksheet
    Dim i As Integer
    Dim endRow As Long
    Dim shtNames As Range
    Dim Cell As Range

    Set mySht = Sheets.Add
    mySht.Move Before:=Sheets(1)
    For i = 2 To Sheets.count
        mySht.Cells(i - 1, 1).Value = Sheets(i).Name
    Next i
    endRow = mySht.Cells(Rows.count, 1).End(xlUp).Row
    Set shtNames = mySht.Range(Cells(1, 1), Cells(endRow, 1))
    shtNames.Sort key1:=Range("A1"), order1:=xlAscending, Header:= _
    xlNo, OrderCustom:=1

    i = 2
    For Each Cell In shtNames
        Sheets(Cell.Value).Move Before:=Sheets(i)
        i = i + 1
    Next Cell

    Application.DisplayAlerts = False
    mySht.Delete
    Application.DisplayAlerts = True
    Call endTimer
    
End Sub


'----------------------------
' @name  IsWorkBookOpen
' @arg   FileName: Takes a filename (in the full form, C:/Documents/FileName.xlsx)
' @descr Tests if this filename is already open in Excel
' @usage Assign a boolean variable to IsWorkBookOpen, can then IF to either Set or Open this WB w/o errors
' @src   http://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open/9373914#9373914
'----------------------------
Function IsWorkBookOpen(FileToTest As String)
    Dim ff As Long, ErrNo As Long
    If Workbooks.CanCheckOut(FileName:=FileToTest) = False Then
        IsWorkBookOpen = False
        Debug.Print FileToTest & " -> Locked file detected!"
    End If
        
    On Error Resume Next
    ff = FreeFile()
    Open FileToTest For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

' Protects all worksheets in the active workbook (cannot edit)
Sub ProtectAll()
    Dim WS As Worksheet
    Dim wb As Worksheets
    Dim count As Integer
    count = ActiveWorkbook.Worksheets.count
    Dim i As Integer
    For i = 1 To count
        Set WS = ActiveWorkbook.Worksheets(i)
        WS.Protect
    Next
End Sub

' Not that helpful - just for testing handling multiple workbooks!
Sub createLoadsOfWorkbooks()

total = 150
For i = 1 To total Step 1
    ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.count)
    
    With Worksheets(Worksheets.count)
        .Range("B:B").ClearContents
        .Range(.Range("A2"), .Cells(.Rows.count, "A").End(xlUp)).SpecialCells( _
                xlCellTypeConstants, 23).Offset(0, 1).FormulaR1C1 = _
                "=RANDBETWEEN(1+50*(ROW()-2),50+50*(ROW()-2))"
        .Range("B:B").Value = .Range("B:B").Value
        .Range("B1").Value = "Random Number"
    End With

Next

End Sub
