'类别=个人常用
'说明=无说明
Sub 统计各数字占比()
On Error Resume Next
Dim all As Range
Set all = Selection
Set all = all.SpecialCells(xlCellTypeVisible)
'-------------
Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(选一个单元格)。", Title:="结果存放", Type:=8)
If tar Is Nothing Then
    Exit Sub
End If
Set tar = tar.Cells(1, 1)
'-------------
Dim r As Range, r1 As Range
Dim cnt As Integer, sum As Double
cnt = 0
sum = 0
For Each r In all.Areas
    sum = sum + WorksheetFunction.sum(r)
Next
'-------------
For Each r In all.Areas
    For Each r1 In r.Cells
        If IsNumeric(r1.Value) Then
            tar.Offset(cnt, 0).Value = r1.Value
            tar.Offset(cnt, 1).Value = CDbl(r1.Value) / sum
            tar.Offset(cnt, 1).Style = "Percent"
            cnt = cnt + 1
        End If
    Next
Next

End Sub
