'类别=
'说明=根据B列最后数据快速合并A列单元格

Sub 按某一列合并格式合并另一列()
    Dim r As Range, n As Integer, beginRow As Integer
    Dim col1 As String
    Dim col2 As String
    Dim maxRow As String
    Dim lastRow As String
    Dim str
    
    str = Application.InputBox("请输入xx列和yy列(以逗号分开)，按xx列合并yy列:", "输入", "A,B")
    If str = False Then Exit Sub
    str = Replace(str, "，", ",")
    arr = Split(str, ",")
    
    
    col2 = arr(0)   '参照列
    col1 = arr(1)  '目标列
    
    
    maxRow = Rows.count
    lastRow = Range(col2 & maxRow).End(xlUp).Row
    
    Range(col1 & "1:" & col1 & lastRow).MergeCells = False  'unmerge
    
    For i = 1 To lastRow
        Set r = Range(col2 & i)
        If r.MergeCells Then
            If r.MergeArea.Columns.count = 1 Then   '合并方向：单列
                If r.MergeArea.Cells.Offset.Address = r.Address Then
                    n = r.MergeArea.count
                    beginRow = r.MergeArea.Cells.Offset.Row
                    Range(col1 & beginRow & ":" & col1 & CStr(beginRow + n - 1)).Merge
                End If
            End If
        End If
    Next i
End Sub
