'���=���˳���
'˵��=��˵��


Sub ���и��ƺ��Ժϲ�()
Dim tar As Range, org As Range
Set org = Selection
If org.Cells.Count >= 65536 Then
    MsgBox "ѡ��ĵ�Ԫ��̫����(����65536��)"
    Exit Sub
End If


Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��", Title:="������", Type:=8)
If tar Is Nothing Then
    Exit Sub
End If
Set tar = tar.Cells(1, 1)

If org.Cells.Count = 1 Then
    tar.Cells(1, 1) = org.Cells(1, 1).Value
    Exit Sub
End If

Dim data, i, j, x, y
data = org
x = 0: y = 0
For i = 1 To UBound(data, 1) '����
    y = 0
    For j = 1 To UBound(data, 2)
        If data(i, j) <> "" Then
            y = y + 1
            tar.Offset(i - 1, y - 1).Value = data(i, j)
        End If
    Next
Next i


End Sub