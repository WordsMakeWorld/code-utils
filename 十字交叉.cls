'放到ThisWorkbook里面，并启用宏
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
 Sh.Cells.Interior.ColorIndex = xlNone
 'Target.EntireColumn.Interior.ColorIndex = 36
 Target.EntireRow.Interior.ColorIndex = 36
End Sub
