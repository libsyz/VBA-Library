
'Using the Find Method
'Find is, by default:
'- Not case sensitive
'- Looks for any matching string

Option Explicit

Sub FindAFilm()

  Range("B3:B15").Find("Ted").Select
  'If the range contains a cell
  'with the value "unexpected", it
  ' will be selected

  Range("B3:B15").Find(What:="Ted", MatchCase:=True).Select
  'This one will perform a case sensitive search'

  Range("B3:B15").Find(What:="Ted", MatchCase:=True, LookAt:=xlWhole).Select

  'This one will perform a case sensitive search
  'and will only look for whole words
  'Detailed find method documentation is on the MS Dev Platform'

End Sub

'Dealing with values not found

'If a value is not found, VBA gives it the value Nothing
'and unfortunately you can't select Nothing.

Sub FindAFilm()

  Dim movieCell as Range

  set movieCell = Range("B3:B15").Find(What:="Ted", MatchCase:=True, LookAt:=xlWhole)

  If movieCell = Nothing then
    msgBox "No movie was found"
  Else
    movieCell.Select
  End If


End Sub


'Using FindNext'

'Find Next is useful when your search query
'could return multiple matches'

'Find Next can only be used if you have already used Find'

'Example Below'

Sub FindAFilm()

  Dim movieCell as Range
  Dim firstFilmCellAddress as string

  Set searchRange = Range("B3:B15")
  Set movieCell = searchRange.Find(What:="Ted", MatchCase:=True, LookAt:=xlWhole)
  If movieCell = Nothing then
    msgBox "No movie was found"
  Else
    firstFilmCell = movieCell.Address
    msgBox "This movie is on cell " & movieCell
    movieCell.Select
    Do
    Set movieCell = searchRange.FindNext(What:="Ted", MatchCase:=True, LookAt:=xlWhole)
    Loop while movieCell.Address <> firstFilmCellAddress
  End If

End Sub

