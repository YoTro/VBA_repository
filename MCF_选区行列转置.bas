'���=����ת��
'˵��=ѡ������ת��
Option Base 1

Sub ѡ������ת��()

    Dim arr(), count
    x = Selection.Rows.count
    y = Selection.Columns.count

    a = Selection.Value
    
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��", Title:="������", Type:=8)
    If tar Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To x    '��
        For j = 1 To y
            tar.Offset(j - 1, i - 1) = a(i, j)
        Next j
    Next i

End Sub


