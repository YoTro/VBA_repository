'类别=新增录入
'说明=无说明

Sub 空白单元格填充固定内容()
    Dim str
    Dim r As Range
    Dim cols, rows
    Dim i, j
    
    str = Application.InputBox("请输入文本内容:", "输入文本内容")
    
    If str = False Then Exit Sub
    

    If Selection.Cells.Count < 1 Then
        MsgBox "请选中一块区域！"
        Exit Sub
    End If
    
    If Selection.Areas.Count > 1 Then Exit Sub
    rows = Selection.Cells.rows.Count
    cols = Selection.Cells.Columns.Count
    
    For i = 1 To rows
        For j = 1 To cols
            Set r = Selection.Cells(i, j)
            If r.Value = "" Then
                r = str
            End If
        Next j
    Next i
End Sub


