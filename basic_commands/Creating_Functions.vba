Option Explicit

'Writing a Simple Function'

Function CustomDate() as String

  CustomDate = Format(date, "dddd dd mmmm yyyy")

End Function

'Calling Functions'

'You can test functions on the immediate window (ctrl+G)'
'?CustomDate

Sub CreateNewSheet ()

  Worksheets.Add
  range("A1").Value = "Created on " & CustomDate '<- Function call here

End Sub

'Adding Parameters'

Function CustomDate (DateToFormat as Date) As String

  CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy")

End Function

'Adding Optional Parameters'

'All Optional Parameters must be placed
'After the compulsory ones

'Example below sets IncludeTime as optional
'And sets the default value to False

Function CustomDate (DateToFormat as Date, Optional IncludeTime as Boolean = False) As String


  If IncludeTime = True Then
    CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy hh:mm:ss")
  Else
    CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy")
  End If

End Function


'### Where are functions available? In which scope?'
