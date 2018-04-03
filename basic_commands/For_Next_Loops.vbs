
'The Syntax of a For Next Loop'

Sub SimpleForLoop ()

  Dim LoopCounter as Integer

  For LoopCounter 1 to 10
    debug.print("we are on iteration number " & LoopCounter)
  Next LoopCounter

End Sub

' You can add Step to the loop to control pace'

Sub ForLoopWithStep ()

  Dim LoopCounter as Integer

  For LoopCounter 1 to 10 Step 2
    debug.print("we are on iteration number " & LoopCounter)
  Next LoopCounter

  'This will print'
  '> we are on interation number 1
  '> we are on iteration number 3
  '> ...

  'step can also be used to count backwards'

    For LoopCounter 10 to 1 Step -1
    debug.print("we are on iteration number " & LoopCounter)
  Next LoopCounter

End Sub

'Exiting from a For Loop'

'Useful if an error is generated or a
'certain condition is met

Sub ExitForLoop ()

  Dim LoopCounter as Integer
  Dim RandomNumber as Double

  For LoopCounter 1 to 10

    RandomNumber = Math.Rnd

    If RandomNumber > 0.2 Then Exit For

  Next LoopCounter

End Sub


'This technique can be used to iterate through
'any given collection

'Practical example to protect all worksheets'

Sub ProtectAllWorksheets ()

  Dim LoopCounter As Integer
  Dim NumberOfWorksheets As Integer

  NumberOfWorksheets = Worksheets.Count

  For LoopCounter to NumberOfWorksheets

    Worksheets(LoopCounter).protect

  Next LoopCounter

End Sub


