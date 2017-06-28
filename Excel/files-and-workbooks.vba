'----------------------------
' @name  CloseAllWorkbooks
' @descr Closes ALL open workbooks without saving
' @usage No parameters, just closes everything without saving
'----------------------------
Sub CloseAllWorkbooks()
For Each wb In Application.Workbooks
    wb.Close False
Next
'Set Application = Nothing
End Sub

