'类别=批注
'说明=批量插入透明批注
Sub 批量插入透明批注()
    On Error Resume Next
    Dim r As Range, msg As String
    msg = InputBox("请输入欲批量插入的透明批注", "提示", "随便输点什么吧")
    If Selection.Cells.Count > 0 Then
        For Each r In Selection
            r.AddComment
            r.Comment.Visible = False
            r.Comment.Text Text:=msg
            r.Comment.Shape.Fill.Visible = msoFalse
        Next
    End If
End Sub



