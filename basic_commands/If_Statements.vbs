
'Simple If Statements'

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  If movieLength > 100 Then
    msgBox movie & " is a quite long movie"
  else
    msgBox movie & " is a short movie"
  End If

End Sub


'ElseIf Statements'
'Additional elseif statements could be added depending on the
'number of conditions to be tested

'You can also nest if/else conditions within an if/else statement'

Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  If movieLength > 100 Then
    msgBox movie & " is a quite long movie"
  ElseIf movieLength > 150 Then
    msgBox movie & " is an unbearable movie"
  Else
    if movie = "Guardians of the Galaxy" Then
      msgBox "This is the Guardians!"
    Else
      msgBox "This is not a Guardians, but at least is a short movie"
  End If

End Sub

'Testing for multiple conditions at once'
'and          -> And
'or           -> Or
'nil          -> Nothing
'diff than    -> <>
'bigger than  -> >
'smaller than -> <


Sub testStringLength

  Dim movie as String
  Dim movieLength as Integer

  movie = "Guardians of the Galaxy"
  movieLength = 140

  If movieLength > 100 Or len(movie) > 15 Then
    msgBox movie & " is a quite long movie"
  else
    msgBox movie & " is a short movie"
  End If

End Sub

'### Look for all control flow operators'
