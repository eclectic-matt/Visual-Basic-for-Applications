'----------------------------
' @name  parseStrToArr
' @descr Splits a string into an array
' @usage Takes a string and delimiter, e.g. "," and returns an array
'----------------------------
Function parseStrToArr(ByRef str As String, delimiter)
Dim Arr() As String
Arr = Split(str, delimiter, Compare:=vbTextCompare)
parseStrToArr = Arr
End Function
