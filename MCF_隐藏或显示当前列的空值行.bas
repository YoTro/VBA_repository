'类别=空值零值
'说明=隐藏或显示当前列的空值行
Sub 隐藏或显示当前列的空值行()
    Dim colName As String
    colName = Split(ActiveCell.Address, "$")(1)

    With Range(colName & ":" & colName).SpecialCells(xlCellTypeBlanks).EntireRow
        .Hidden = Not .Hidden
    End With
End Sub




