'类别=行列转换
'说明=选区行列转置
Option Base 1

Sub 选区行列转置()

    Dim arr(), count
    x = Selection.Rows.count
    y = Selection.Columns.count

    a = Selection.Value
    
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格。", Title:="结果存放", Type:=8)
    If tar Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To x    '行
        For j = 1 To y
            tar.Offset(j - 1, i - 1) = a(i, j)
        Next j
    Next i

End Sub


