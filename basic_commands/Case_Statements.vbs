'case_statements
'Awesome to make multiple
'if/else conditions more readable and
'maintainable
Option Explicit

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  Select Case movieLength
    Case is < 100
      msgBox "The movie is short"
    Case is > 100
      msgBox "The movie is long"
    Case Else
      msgBox "The movie is strange"
  End Select

End Sub

'Testing for a range of numbers'

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  Select Case movieLength
    Case is 0 To 100
      msgBox "The movie is short"
    Case is 100 To 200
      msgBox "The movie is long"
    Case Else
      msgBox "The movie is strange"
  End Select

End Sub


'Testing for several values'

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  'Case will evaluate to true is MovieLength
  'Is on any of the comma-separated value lists
  Select Case movieLength
    Case is 80, 90, 100
      msgBox "The movie is short"
    Case is 100, 120, 140
      msgBox "The movie is long"
    Case Else
      msgBox "The movie is strange"
  End Select

End Sub

' Case Statements can also be nested'

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  Select Case movieLength
    Case is 80, 90, 100
      Select Case movieLength
      Case is 80 to 89
        msgBox "The length of this movie is under 90"
      Case is 90 to 99
        msgBox "The length of this movie is under 100"
      Case Else
        msgBox "The length of this movie is 100"
      End Select
    Case is 100, 120, 140
      msgBox "The movie is long"
    Case Else
      msgBox "The movie is strange"
  End Select

End Sub



