'类别=个人常用
'说明=计算选区(必须为单列)连续相同值的单元格数目
Sub 连续相同值的个数统计()

Dim r As Range, tmpr As Range
Dim total As Integer

    If Selection.Columns.Count > 1 Then
        MsgBox "选区只允许包含一个列！"
        Exit Sub
    End If

Set tmpr = Nothing

For Each r In Selection
    If tmpr Is Nothing Then
        Set tmpr = r
        total = 1
    Else
        If r.Value = tmpr.Value Then  '一样
            total = total + 1
        Else  '不一样
            tmpr.Offset(0, 1) = total
            Set tmpr = r
            total = 1
        End If
    End If
Next

If Not tmpr Is Nothing Then tmpr.Offset(0, 1) = total

End Sub



