'类别=空值零值
'说明=删除选区单列的空值单元格整行

Sub 删除选区的空值单元格整行()
    Dim colName As String

    With Selection.SpecialCells(xlCellTypeBlanks).EntireRow
        .Delete
    End With
End Sub






