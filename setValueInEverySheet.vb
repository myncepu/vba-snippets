Sub setValueInEverySheet()
'
' setValuesInEverySheet Macro
' 宏由 Administrator 录制，时间: 2020/05/25
' 设置每个 sheet B4 值，依次为 2020-5-25, 2020-5-26, 2020-5-27...
'

'
   
    Dim sheetIdx, sheetCount As Integer

    For sheetIdx = 1 To Sheets.Count
      ' DATEVALUE("2020-5-25") = 43976
      Sheets(sheetIdx).Cells(4, 2).Value = CDate(43976 + sheetIdx)
      ' MsgBox (d)
    Next
    
End Sub