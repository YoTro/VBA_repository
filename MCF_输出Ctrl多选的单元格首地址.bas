'���=��λ����
'˵��=��˵��


Sub ���Ctrl��ѡ�ĵ�Ԫ���׵�ַ()
On Error Resume Next
Dim r As Range, tar As Range

Set tar = Application.InputBox("ѡ����λ��", Type:=8)
If tar Is Nothing Then Exit Sub


tar.Offset(0, 0) = "��Ԫ��"
tar.Offset(0, 1) = "�к�"
tar.Offset(0, 2) = "�к�"
Dim cnt As Integer
cnt = 1
For Each r In Selection.Areas
    Set r = r.Cells(1, 1)
    tar.Offset(cnt, 0) = r.Address(False, False)
    tar.Offset(cnt, 1) = r.Row
    tar.Offset(cnt, 2) = r.Column
    cnt = cnt + 1
Next

End Sub