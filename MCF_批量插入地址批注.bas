'类别=批注
'说明=批量插入地址批注
Sub 批量插入地址批注()
    On Error Resume Next
    Dim r As Range
    If Selection.Cells.Count > 0 Then
        For Each r In Selection
            r.Comment.Delete
            r.AddComment
            r.Comment.Visible = False
            r.Comment.Text Text:="本单元格：" & r.Address & " of " & Selection.Address
        Next
    End If
End Sub




