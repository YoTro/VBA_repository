'类别=数值转换
'说明=无说明
Sub 修复不完整的换行符()
Dim r As Range
Dim s As String

For Each r In Selection.Cells
    s = r.Value
    's = Replace(s, vbCr, vbCrLf)
    s = Replace(s, vbLf, vbCrLf)
    r.Value = s
Next
msgbox "完成"
End Sub
