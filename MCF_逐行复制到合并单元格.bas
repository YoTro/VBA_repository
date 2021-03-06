'类别=个人常用
'说明=无说明

Sub 逐行复制到合并单元格()
Dim tar As Range, org As Range
Set org = Selection
If org.Cells.Count >= 65536 Then
    MsgBox "选择的单元格太大了(超过65536个)"
    Exit Sub
End If


Set tar = Application.InputBox(prompt:="请选择存放结果的单元格", Title:="结果存放", Type:=8)
If tar Is Nothing Then
    Exit Sub
End If
Set tar = tar.Cells(1, 1)

If org.Cells.Count = 1 Then
    tar.Cells(1, 1) = org.Cells(1, 1).Value
    Exit Sub
End If

Dim r As Range
Dim data, i, j
data = org
For i = 1 To UBound(data, 1) '逐行
    Set r = tar.Offset(i - 1, 0)
    For j = 1 To UBound(data, 2)
        Set r = GetRightUnMergeRange(r)
        If Not r Is Nothing Then
            r.Value = data(i, j)
            Set r = r.Offset(0, 1)
        End If
    Next
Next i





End Sub

Function GetRightUnMergeRange(tar As Range) As Range
On Error Resume Next
Dim r As Range

For i = 0 To Rows.Count
    Set r = tar.Offset(0, i)
    
    If r.MergeCells Then '合并
        If r.MergeArea.Cells.Offset.Address = r.Address Then  '首个
            Set GetRightUnMergeRange = r
            Exit Function
        End If
    Else '非合并
        Set GetRightUnMergeRange = r
        Exit Function
    End If
Next

Set GetRightUnMergeRange = Nothing
End Function