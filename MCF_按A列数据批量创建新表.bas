'类别=工作表
'说明=按A列数据批量创建新表(控件按钮代码)

Private Sub 按A列数据批量创建新表()
On Error Resume Next
Dim i%, j%

For i = 1 To [a65536].End(xlUp).Row
    For j = 2 To Sheets.Count
        If Cells(i, 1) = Sheets(j).Name Then
            Exit For
        End If
    Next
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = Cells(i, 1)
Next

End Sub


