' Functions to set up timers and progress bars in VBA
' NOTE: Timer functions could be used in other applications with VBA

'----------------------------
' @name  startTimer
' @descr STARTS the timer to test macro run times
' @usage CALL at the beginning of a module and then CALL endTimer before "End Sub"
'----------------------------
Sub startTimer()
  tStart = Timer
End Sub

'----------------------------
' @name  endTimer
' @descr ENDS the timer and displays macro run time
' @usage CALL before "End Sub" - make sure startTimer called at beginning!
'----------------------------
Sub endTimer()
  timeTaken = Timer - tStart
  'tMins = Round(timeTaken / 60, 0)
  tMins = Int(timeTaken / 60)
  tSecs = Format(timeTaken Mod 60, "00")

  'tMins = timeTaken Mod 60
  'tSecs = timeTaken - (60 * tMins)
  MsgBox ("This process took " & timeTaken & " secs" & _
          vbNewLine & vbNewLine & _
          tMins & " minutes and " & tSecs & " seconds." & _
          vbCr & tMins & ":" & tSecs)
End Sub


Sub ProgBar(total)
  If Complete = True Then GoTo Completed:
  tStart = Timer

  ' Local vars - passed to ProgressBar subs as "num" and "den"
  Dim Numerator As Integer
  Dim Denominator As Integer
  Dim PctDone As Single

  Application.ScreenUpdating = False

  TimeLeft = ""

  ' Loop through cells.
  For r = 1 To RowMax

      For c = 1 To ColMax
          Numerator = Numerator + 1
      Next c

      ' Update the percentage completed.
      PctDone = Numerator / Denominator

      ' Call subroutine that updates the progress bar.
      Call UpdateProgressBar(PctDone, Numerator, Denominator)

  Next r

  Completed:

  Complete = True
  With ProgressForm
      .FrameProgress.Caption = "COMPLETED."
      .LabelProgress.Caption = "TASK COMPLETE."
      .Label2.Caption = "SUCCESS - " & Denominator & " records updated."
      .Label1.Caption = "Time taken: approx. " & Round(t, 1) & " seconds."
  End With
  ProgressForm.Show
  t = Timer - tStart
  MsgBox ("Just Completed in " & t & " secs")

  'tStart = Timer
  't = Timer - tStart
  'Do While t < timeOut
  '    ProgressForm.LabelProgress.Caption = "CLOSING IN " & Round(timeOut - t, 1) & " SECONDS"
  '    'DoEvents
  '    t = Timer - tStart
  'Loop

  'ProgressForm.Hide
  Call Unload(ProgressForm)

  'MsgBox ("Finally Completed.")
End Sub
