'类别=合并和拆分
'说明=按选定列的合并格式合并其他列
Sub 应用选区的合并格式到其他列()
    On Error GoTo l_err
    
    Dim r As Range
    Dim i, n As Integer, beginRow As Integer
    Dim cols As String
    Dim arr() As String, colTgt As String
    
    If Selection.Columns.Count > 1 Then
        MsgBox "选区不允许包含多个列！"
        Exit Sub
    End If

    cols = Application.InputBox(prompt:="输入要合并的列名(用逗号隔开，如 E,F,G):", Type:=2,Default:="I,J,K,M,N,O")
    arr = Split(cols, ",")
    
    Application.DisplayAlerts = False
    
    For Each r In Selection
        If r.MergeCells Then
            If r.MergeArea.Columns.Count = 1 Then   '合并方向：单列
                If r.MergeArea.Cells.Offset.Address = r.Address Then
                    n = r.MergeArea.Count
                    beginRow = r.MergeArea.Cells.Offset.Row
                    
                    For i = 0 To UBound(arr)
                        colTgt = arr(i)
                        Range(colTgt & beginRow & ":" & colTgt & CStr(beginRow + n - 1)).Merge
                    Next i
                End If
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
    Exit Sub
l_err:
    Application.DisplayAlerts = True
    MsgBox "发生错误：" & Err.Description
End Sub










