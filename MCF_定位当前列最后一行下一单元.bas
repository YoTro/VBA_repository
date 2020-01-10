'类别=定位引用
'说明=光标定位到指定工作表A列最后数据行下一单元

Sub 定位当前列最后一行下一单元()
    Dim colName As String
    Dim maxRow As String
    dim lastRow as String
    
    colName = Split(ActiveCell.Address, "$")(1)
    maxRow = Rows.Count

    lastRow = Range(colName & maxRow).End(xlUp).Row
    'lastRow = ActiveSheet.[a65536].End(xlUp).Row

    Range(colName  & lastRow + 1).Select
    
End Sub





