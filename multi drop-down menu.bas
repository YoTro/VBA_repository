Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'多项选择的下拉菜单
    If Target.Column <> 2 Or Target.Row < 4 Then ListBox1.Visible = False: Exit Sub
    '如果选中的单元格不是第2列，或者小于4行，也就是不在目标范围内，则退出程序
    If Target.Columns.Count > 1 Or Target.Rows.Count > 1 Then ListBox1.Visible = False: Exit Sub
    '如果选中的单元格大于1个，则退出程序
    With Sheets("参数表")
        r = .Range("a1:c" & .Cells(Rows.Count, "a").End(xlUp).Row).Value
    End With
    With ListBox1
        '调整位置到单元格处
        .Top = Target.Top 'listbox的顶端位置
        .Left = Target.Left + Target.Width 'listbox的左端位置
        .Width = 250 '宽度
        .Height = 150 '高度
        .Visible = True '可见
        '.ColumnHeads = True '显示标题行
        .ColumnCount = 3 '三列
        .ColumnWidths = "50;120;50" '设置第一列宽度50第二列宽度120……
        .List = r '数据来源
        .MultiSelect = fmMultiSelectMulti '允许通过鼠标点击的方式进行多选
        .ListStyle = fmListStyleOption '选项按钮设置为方形
    End With
End Sub

Private Sub ListBox1_Change()
    Dim i As Long, strMy As String
    With ListBox1
        If .Selected(0) = True Then .Selected(0) = False
        '如果用户选取的是标题行那么撤销选取
        For i = 1 To .ListCount - 1
        '遍历listbox的记录，如果被选中则按换行符合并
            If .Selected(i) = True Then
                strMy = strMy & vbCrLf & .List(i, 1)
                '取list的第二列
                '无论列还是行的索引都是从0开始的，因此第二列为1
            End If
        Next
    End With
    ActiveCell.Value = Mid(strMy, 3)
    '数据写入单元格
End Sub