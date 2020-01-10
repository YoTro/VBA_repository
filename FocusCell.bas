Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'聚焦单元格所在行列填充颜色

    Application.ScreenUpdating = False

        Cells.Interior.ColorIndex = -4142

        '取消单元格原有填充色，但不包含条件格式产生的颜色。

        Rows(Target.Row).Interior.ColorIndex = 33

        '活动单元格整行填充颜色

        Columns(Target.Column).Interior.ColorIndex = 33

        '活动单元格整列填充颜色

    Application.ScreenUpdating = True

End Sub