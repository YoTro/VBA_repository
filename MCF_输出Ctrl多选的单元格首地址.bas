'类别=定位引用
'说明=无说明


Sub 输出Ctrl多选的单元格首地址()
On Error Resume Next
Dim r As Range, tar As Range

Set tar = Application.InputBox("选择存放位置", Type:=8)
If tar Is Nothing Then Exit Sub


tar.Offset(0, 0) = "单元格"
tar.Offset(0, 1) = "行号"
tar.Offset(0, 2) = "列号"
Dim cnt As Integer
cnt = 1
For Each r In Selection.Areas
    Set r = r.Cells(1, 1)
    tar.Offset(cnt, 0) = r.Address(False, False)
    tar.Offset(cnt, 1) = r.Row
    tar.Offset(cnt, 2) = r.Column
    cnt = cnt + 1
Next

End Sub