'类别=空值零值
'说明=删除当前列的空值单元格整行

Sub 删除当前列的空值单元格整行()
    Dim colName As String
    colName = Split(ActiveCell.Address, "$")(1)

    With Range(colName & ":" & colName).SpecialCells(xlCellTypeBlanks).EntireRow
        .Delete
    End With
End Sub






