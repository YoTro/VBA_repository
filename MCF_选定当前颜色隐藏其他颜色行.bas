'类别=定位引用
'说明=选定当前颜色隐藏其他颜色整行
Sub 选定当前颜色隐藏其他颜色行() '思路就是:其它背景色之行全部隐藏
    Dim UseRow, AC, i '首先选择一个有颜色之单元格，然后动行宏，其它颜色所在行隐藏
    UseRow = Cells.SpecialCells(xlCellTypeLastCell).Row 'SpecialCells(xlCellTypeLastCell)表示已用区域最后一个单元格
    If ActiveCell.Row > UseRow Then
        MsgBox "请在要筛选的区域选择一个有颜色之单元格！", vbExclamation, "错误"
    Else
        AC = ActiveCell.Column
        Cells.EntireRow.Hidden = False '显示所有行
        
        For i = 2 To UseRow
            If Cells(i, AC).Interior.ColorIndex <> ActiveCell.Interior.ColorIndex Then
                Cells(i, AC).EntireRow.Hidden = True '如果2至已用行之单元格的有列之颜色不等于当前单元格颜色则隐藏整行
            End If
        Next
    End If
End Sub



