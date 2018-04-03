
'Simple Do Loop'

Sub SimpleDoLoop()
  Range("A1").Select

  'Infinite Loop below - no break condition'
  Do
    ActiveCell.Offset(0,1)
  Loop

End Sub


'A condition needs to stop the loop
'from running either with While or Until

Sub UntilDoLoop()
  Range("A1").Select

  'Loop stops on an Empty Cell'
  Do Until ActiveCell.Value = ""
    ActiveCell.Offset(0,1)
  Loop

  'Until condition can also be applied
  'On the closing line of the loop
  'thus making sure your loop runs at least once'
  Do
    ActiveCell.Offset(0,1)
  Loop Until ActiveCell.Value = ""

End Sub

Sub WhileDoLoop()

  Application.ScreenUpdating = False 'You only see the end result!'

  Range("A1").Select

  'Loop stops on an Empty Cell'
  Do While ActiveCell.Value <> ""
    ActiveCell.Offset(0,1)
  Loop

  'While condition can also be applied
  'On the closing line of the loop
  'thus making sure your loop runs at least once'
  Do
    ActiveCell.Offset(0,1)
  Loop While ActiveCell.Value <> ""

  Application.ScreenUpdating = True

End Sub

'Using Exit to break from a Do Loop'

Sub ExitDoLoop()
  Range("A1").Select

  Do
    If ActiveCell.Value = "" Then Exit Do
    ActiveCell.Offset(0,1)
  Loop

  'The advantage of this technique is that
  'the stopping condition can be placed
  'at any point in time

  Dim counter as Integer

  Do
    counter = counter + 1
    If counter = 50 Then Exit Do
    ActiveCell.Offset(0,1)
  Loop

End Sub






