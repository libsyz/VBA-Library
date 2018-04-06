
'Awesome for looping over collections'

'Looping over Worksheets'

Sub listWorksheetNames ()

  dim singleSheet as workSheet

  For Each singleSheet in WorkSheets
    debug.print(singleSheet.name)
  Next  singleSheet

End Sub

'Looping over Workbooks'

Sub closeAllWorkbooks ()

  dim singleWorkbook as workbook

  For Each singleWorkbook in Workbooks
    singleWorkbook.close
  Next  singleWorkbook

End Sub

'Qualifying Collection Names'

'Particularly useful if you are processing
'chart objects

Sub ProtectSheetsInAnotherWorkbook ()

  Dim singleSheet as Worksheet

  For Each singleSheet in Workbooks("Book2.xlsx").WorkSheets
    singleSheet.Protect
  Next singleSheet

End Sub

'Looping over a Range of Cells

Sub LoopingOverARangeOfCells ()

  dim SingleCell as Range

  For Each SingleCell in Range("RangeName")
    debug.print SingleCell.Value
  Next SingleCell

End Sub

'Nested For Each Loops '

Sub nestedForEachLoop ()

  Dim singleWorkbook as Workbook
  Dim singleSheet as WorkSheet

  For Each singleWorkbook in Workbooks
    For Each singleSheet in singleWorkbook
      singleSheet.protect
    Next singleSheet
  Next Workbook

End Sub




