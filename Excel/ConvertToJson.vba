Sub ConvertToJson()
    
    'Select the empty column you want to populate with JSON
    endCol = Selection.Cells(1, Selection.Columns.Count).Column
    'Calculate the end row of data
    endRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'Iterate through the rows in the sheet
    For Row = 2 To endRow Step 1
        
        'Initialise output
        output = "{"
        
        'Iterate through columns (up to endCol - 1)
        For Column = 1 To (endCol - 1) Step 1
        
            'Cells(5, 3) = C5 = row 5, column 3
            'Get the header/title (assumed in top row)
            Heading = ActiveSheet.Cells(1, Column).value
            'Get the current value
            value = ActiveSheet.Cells(Row, Column).value
            'Generate a JSON string (note: no formatting or validity checks)
            output = output & Chr(34) & Heading & Chr(34) & ":" & Chr(34) & value & Chr(34) & ","
            
        'Next column
        Next Column
        
        'Once all columns processed, close output string
        output = output & "},"
        'And store output string in the end column
        ActiveSheet.Cells(Row, endCol).value = output
        
    'Next row
    Next Row
        
End Sub
