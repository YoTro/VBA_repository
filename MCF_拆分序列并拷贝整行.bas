'类别=个人常用
'说明=无说明


Sub 拆分序列并拷贝整行()
'3-5 可拆分为 3、4、5
Dim tar As Range, mo As Range

Set tar = Selection
If tar.Columns.Count <> 1 Then
    MsgBox "只能选择单列"
    Exit Sub
End If

Set mo = Application.InputBox(prompt:="请选择存放结果的单元格(将整行复制)", Title:="结果存放", Type:=8)
If mo Is Nothing Then
    Exit Sub
End If
Set mo = mo.Cells(1, 1)
'---------------------------
Dim arr, tmp, r As Range
Dim x1 As Integer, x2 As Integer, cnt As Integer
cnt = 0
Dim col As Integer

For Each r In tar.Cells
    tmp = r.Value
    arr = Split(tmp, "-")
    If UBound(arr) >= 1 Then
            x1 = arr(0)
            x2 = arr(1)
            
            col = r.Column
            '------------------
            For i = x1 + 1 To x2
                cnt = cnt + 1
                r.EntireRow.Copy mo.Offset(cnt).EntireRow
                mo.Offset(cnt).EntireRow.Cells(1, col) = i
            Next
            r.Value = x1
    End If
Next
End Sub