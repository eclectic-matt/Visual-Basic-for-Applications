'----------------------------
' @name  getStdDevSam
' @descr An attempt to avoid the WORKSHEET FUNCTION version of StdDev.S
' @usage Takes a range and returns the SAMPLE standard deviation
' @form  Math.SQR( Sum(Value - Mean)^2 / (SampleSize - 1) )
'----------------------------
Function getStdDevSam(ByVal stdRng As Range)
    numItems = stdRng.count
    mean = getAverage(stdRng)
    If mean = "N/A" Then
        deviation = "N/A"
    Else
        devSum = 0
        For Each Cell In stdRng
            If Not Len(Cell) = 0 Then
                devSum = devSum + (Cell - mean) ^ 2
            Else
                numItems = numItems - 1
            End If
        Next
        deviation = Math.Sqr((devSum / (numItems - 1)))
    End If
    getStdDevSam = deviation
End Function

'----------------------------
' @name  getStdDevPop
' @descr An attempt to avoid the WORKSHEET FUNCTION version of StdDev.P
' @usage Takes a range and returns the POPULATION standard deviation
' @form  Math.SQR( Sum(Value - Mean)^2 / (PopulationSize) )
'----------------------------
Function getStdDevPop(ByVal stdRng As Range)
    numItems = stdRng.count
    mean = getAverage(stdRng)
    If mean = "N/A" Then
        deviation = "N/A"
    Else
        devSum = 0
        For Each Cell In stdRng
            If Not Len(Cell) = 0 Then
                devSum = devSum + (Cell - mean) ^ 2
            Else
                numItems = numItems - 1
            End If
        Next
        deviation = Math.Sqr((devSum / numItems))
    End If
    getStdDevPop = deviation
End Function

'----------------------------
' @name  getMedian
' @descr An attempt to avoid the WORKSHEET FUNCTION version of Median
' @usage Takes a range and returns the median value (sorted as range as no VBA Arr Sort)
'----------------------------
Function getMedian(ByRef medRng As Range)
    medRng.Sort key1:=medRng, order1:=xlAscending, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, Header:=xlNo, DataOption1:=xlSortTextAsNumbers
    Dim numItems As Integer
    'TESTING COUNT
    'Debug.Print "Rng Cnt: " & medRng.count
    'Debug.Print "EA.CntA: " & Application.WorksheetFunction.CountA(medRng)
    'Debug.Print "Blk Cnt: " & countIn(medRng, "")
    numItems = medRng.count - countIn(medRng, "")
    If numItems = 0 Then
        midAvg = "N/A"
    Else
        If numItems = 1 Then
            midAvg = medRng.Value2
        Else
            Dim Arr() As Variant
            Arr = medRng
            If numItems Mod 2 = 0 Then
                middle1 = Arr(Round(numItems / 2, 0), 1)
                middle2 = Arr(Round(numItems / 2, 0) + 1, 1)
                midAvg = (middle1 + middle2) / 2
            Else
                midAvg = Arr(Round(numItems / 2, 0), 1)
            End If
        End If
    End If
    getMedian = midAvg
End Function

'----------------------------
' @name  getAverage
' @descr An attempt to avoid the WORKSHEET FUNCTION version of Average
' @usage Takes a range and returns the arithmetic average (mean) value
'----------------------------
Function getAverage(ByRef avgRng As Range)
    numItems = avgRng.count
    avgSum = 0
    For Each Cell In avgRng
        If IsNumeric(Cell) And Not Len(Cell) = 0 Then
            avgSum = avgSum + Cell
        Else
            numItems = numItems - 1
        End If
    Next
    If numItems = 0 Then
        avg = "N/A"
    Else
        avg = avgSum / numItems
    End If
    getAverage = avg
End Function

'----------------------------
' @name  getMax
' @descr An attempt to avoid the WORKSHEET FUNCTION version of MAX
' @usage Takes a range and returns the largest number found (converts strings to DOUBLES)
'----------------------------
Function getMax(ByRef maxRng As Range)
Dim maxFound
maxFound = MinDouble
For Each Cell In maxRng
    If IsNumeric(Cell) And Len(Cell) <> 0 Then
        NumToCheck = CDbl(Cell)
        If NumToCheck > maxFound Then
            maxFound = NumToCheck
        End If
    End If
Next
If maxFound = MinDouble And countIn(maxRng, "0") = 0 Then
    getMax = "N/A"
Else
    getMax = maxFound
End If
End Function

' MaxDouble and MinDouble from: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1003%20MinMaxVals.shtml
Function MinDouble() As Double
    MinDouble = -1.79769313486231E+308
End Function

Function MaxDouble() As Double
    MaxDouble = 1.79769313486231E+308
End Function


