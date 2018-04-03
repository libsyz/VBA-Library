
'with are great when you have to call
' different methods and properties
' on the same object

sub FormatCells ()

  With Range("C5", Range("C5").End(xlDown))
      .Interior.Color = rgbAquaMarine
      .Font.Color = rbgRed
  end With

  'References to other objects also fit on with statements'

end Sub


