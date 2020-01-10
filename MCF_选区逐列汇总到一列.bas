'类别=行列转换
'说明=选区逐列汇总到一列
Option Base 1

Sub 选区逐列汇总到一列()

    Dim arr(), count
    x = Selection.Rows.count
    y = Selection.Columns.count

    a = Selection.Value
    
    count = 0
    ReDim arr(1 To Selection.count)
    For i = 1 To y    '优先按列
        For j = 1 To x
            count = count + 1
            arr(count) = a(j, i)
        Next j
    Next i
    
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(存放不重复序列,按列)。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(count, 1) = WorksheetFunction.Transpose(arr)  '按列写入
    'tar.Resize(1, count) = WorksheetFunction.Transpose(arr)  '按行写入
End Sub



