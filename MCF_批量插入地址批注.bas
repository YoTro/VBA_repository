'���=��ע
'˵��=���������ַ��ע
Sub ���������ַ��ע()
    On Error Resume Next
    Dim r As Range
    If Selection.Cells.Count > 0 Then
        For Each r In Selection
            r.Comment.Delete
            r.AddComment
            r.Comment.Visible = False
            r.Comment.Text Text:="����Ԫ��" & r.Address & " of " & Selection.Address
        Next
    End If
End Sub




