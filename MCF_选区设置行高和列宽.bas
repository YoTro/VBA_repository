'类别=
'说明=选区行高和列宽
Sub 选区设置行高和列宽()
    Dim str, arr
    
    str = Application.InputBox("请输入行高和列宽，以逗号分开:", "输入", "10,12")
    If str = False Then Exit Sub
    
    str = Replace(str, "，", ",")
    
    arr = Split(str, ",")
    Selection.RowHeight = CInt(arr(0))   '指定行高
    Selection.ColumnWidth = CInt(arr(1))  '指定列宽
    
End Sub








