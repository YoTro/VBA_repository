'类别=定位引用
'说明=选择下一行




Sub 选择下一行()
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
End Sub



