'类别=个人常用
'说明=无说明
Sub 只保留显示的值()
On Error Resume Next
Dim r As Range
Dim all As Range
Set all = Application.Intersect(Selection, Selection.Worksheet.UsedRange)

For Each r In all.Cells
    If r.Value <> "" Then
        txt = r.Text
        r.NumberFormatLocal = "@"
        r.Value = txt
    End If
Next
End Sub
