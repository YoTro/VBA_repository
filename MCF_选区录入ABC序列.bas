'类别=新增录入
'说明=选区录入ABC序列


Sub 选区录入ABC序列()
    Dim i  As Integer
    'Selection.ClearContents '清除内容
    
    i = 0
    For Each Rng In Selection
        If Rng.MergeCells Then  '是否是合并单元格
             If Rng.MergeArea.Cells.Offset.Address = Rng.Address Then  '是否是合并单元格的第一个
                Rng.Value = Chr(65 + i)
                i = (i + 1) Mod 26
             End If
        Else
            Rng.Value = Chr(65 + i)
            i = (i + 1) Mod 26
        End If
    Next
    
End Sub







