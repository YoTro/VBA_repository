'类别=数值转换
'说明=无说明
Sub 两区域交换内容()
On Error Resume Next
Dim tar As Range
Set tar = Selection
'-------------
Dim t1 As Range, t2 As Range
If tar.Areas.count = 1 Then
    MsgBox "请先选中2个区域（提示：按住Ctrl选择多个区域）"
    Exit Sub
ElseIf tar.Areas.count = 2 Then
    Set t1 = tar.Areas(1)
    Set t2 = tar.Areas(2)
    If t1.Rows.count <> t2.Rows.count Or t1.Columns.count <> t2.Columns.count Then
        MsgBox "2个区域大小不一致！"
        Exit Sub
    End If
Else
    MsgBox "请先选中2个区域"
    Exit Sub
End If
'-----------
Dim tmp1, tmp2
tmp1 = t1.Value
tmp2 = t2.Value
t1.Value = tmp2
t2.Value = tmp1

End Sub