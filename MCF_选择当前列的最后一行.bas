'类别=定位引用
'说明=选择当前列的最后一行
Sub 选择当前列的最后一行()
    Dim colName As String
    Dim maxRow As String
    
    colName = Split(ActiveCell.Address, "$")(1)
    maxRow = Rows.Count
    Range(colName & maxRow).End(xlUp).EntireRow.Select
    
    
    'ActiveCell.End(xlDown).EntireRow.Select
    'ActiveCell.SpecialCells(xlCellTypeLastCell).EntireRow.Select
End Sub



