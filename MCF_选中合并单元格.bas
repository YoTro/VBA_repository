'类别=合并和拆分
'说明=选中合并的单元格
Sub 选中合并单元格()
    On Error GoTo l_err
    
    Dim r As Range, allUsed As Range
    Dim all As Range

    Set allUsed = Intersect(ActiveSheet.UsedRange, Selection)
    
    For Each r In allUsed
        If r.MergeCells Then
        '------------------------------
            If all Is Nothing Then
                Set all = r
            Else
                Set all = Union(all, r)
            End If
        '------------------------------
        End If
    Next
    
     If Not all Is Nothing Then all.Select
    Exit Sub
l_err:
    MsgBox "发生错误：" & Err.Description
End Sub

