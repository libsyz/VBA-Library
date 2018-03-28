
'Normal inputs are quite limited in scope
' - You can't control for data input quality
' - You can't use the spreadsheet while an input window is active

Sub ApplicationInputBoxExamples ()

  ' Application.InputBox is different from InputBox!'
  filmName = Application.InputBox("Please enter a film name")

  'You can specify which data type is returned from the InputBox'
  ' The line below has input validation built in already'
  ' This box also allows for clicking outside of the dialog box
  filmLength = Application.InputBox(Prompt:="Please enter a length",
                                    Type:= 1)

  'Entering dates requires a workaround since Application.InputBox
  'does not have a built in date type validation

  filmDate = Application.InputBox(Prompt:="Please enter a date dd/mm/yy", _
                                    Type:= 1)

End Sub


sub EnterFormulaIntoCell ()

  Dim MyFormula as String

  MyFormula = Application.InputBox(Prompt:="Please enter a date dd/mm/yy", _
                                    Type:= 0)

  range("G2").formulaLocal = MyFormula
  'You can also add default values'
End Sub


sub EnterFormulaIntoRange ()
  Dim MyRange as Range

  Set MyRange = Application.InputBox(Prompt:="Enter a range", _
                                     Type:=8)

  MyRange.formulaLocal = MyFormula

End Sub


Sub CopyData ()

    Dim CopyRange As Range
    Dim DestinationRange as Range

    Set CopyRange = Application.InputBox(Prompt:="Enter a range", _
                                          Type:=8)

    Set DestinationRange = Application.InputBox(Prompt:= "Click on a destination range", _
                                                Type:=8)

    CopyRange.Copy DestinationRange

End Sub


Sub ReturnArrayValuesFromRange ()

  Dim FilmLengths() as Variant

  FilmLengths = Application.InputBox(Prompt:="Enter a range to convert", _
                                    Type:=64)

End Sub

