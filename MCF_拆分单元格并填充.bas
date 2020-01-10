'类别=合并和拆分
'说明=拆分合并的单元格并填充

Sub 拆分单元格并填充()
    On Error GoTo l_err
    
    Dim r As Range
    Dim rt As Integer, ct As Integer
    Dim i, j
    Dim tmpV

    
    For Each r In Selection
        If r.MergeCells Then
        '------------------------------
            If r.MergeArea.Cells.Offset.Address = r.Address Then
                tmpV = r.Value
                rt = r.MergeArea.Rows.Count
                ct = r.MergeArea.Columns.Count
                '-----------------------------
                r.UnMerge
                For i = 0 To rt - 1
                    For j = 0 To ct - 1
                        r.Offset(i, j) = tmpV
                    Next j
                Next i
                
            End If
        '------------------------------
        End If
    Next
    Exit Sub
l_err:
    MsgBox "发生错误：" & Err.Description
End Sub



