'���=��ע
'˵��=��������ͳһ��ע

Sub ��������ͳһ��ע()
    On Error Resume Next
    Dim r As Range, msg As String
    msg = InputBox("�������������������ע", "��ʾ", "������ʲô��")
    If Selection.Cells.Count > 0 Then
        For Each r In Selection
            r.AddComment
            r.Comment.Visible = False
            r.Comment.Text Text:=msg
        Next
    End If
End Sub





