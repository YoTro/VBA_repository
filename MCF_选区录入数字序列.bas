'类别=新增录入
'说明=选区录入数字序列，如1,2,3,4...
Sub 选区录入数字序列()
    Dim i  As Integer
    'Selection.ClearContents '清除内容
    
    i = 0
    For Each Rng In Selection
    
        If Rng.MergeCells Then  '是否是合并单元格
             If Rng.MergeArea.Cells.Offset.Address = Rng.Address Then  '是否是合并单元格的第一个
                i = i + 1
                Rng.Value = i
             End If
        Else
            i = i + 1
            Rng.Value = i
        End If
    Next
    
End Sub








