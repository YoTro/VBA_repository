'类别=定位引用
'说明=按指定文本定位

Sub 定位文本()
    Dim r As Range, a As Range
    Dim s
    s = Application.InputBox("请输入要定位的文本:", "输入要定位的文本:", "合计")
    
    If s = False Then Exit Sub
    
    For Each a In ActiveSheet.UsedRange
        If a Like "*" & s & "*" Then
            If r Is Nothing Then
                Set r = a.Cells
            Else
                Set r = Union(r, a.Cells)
            End If
        End If
    Next
    
    If r Is Nothing Then
        MsgBox "未找到指定单元格!"
    Else
        r.Select
    End If
End Sub






