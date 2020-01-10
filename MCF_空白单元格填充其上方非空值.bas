'类别=新增录入
'说明=
Sub 空白单元格填充其上方非空值()
    Dim r As Range, tmp
    Dim cols, rows
    Dim i, j
    
    If Selection.Cells.Count <= 1 Then
        MsgBox "请选中一块区域！"
        Exit Sub
    End If
    
    If Selection.Areas.Count > 1 Then Exit Sub
    rows = Selection.Cells.rows.Count
    cols = Selection.Cells.Columns.Count
    
    For j = 1 To cols
        tmp = ""
            For i = 1 To rows
            Set r = Selection.Cells(i, j)
            If r.Value = "" Then
                r = tmp
            Else
                tmp = r
            End If
        Next i
    Next j
End Sub
