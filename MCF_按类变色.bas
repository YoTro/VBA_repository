'���=��λ����
'˵��=����Ԫ�����ɫ
Sub �����ɫ()

Dim rng As Range

For Each rng In Selection
If rng < 0 Then
rng.Interior.ColorIndex = 4   'С��0�ĵ�Ԫ���̵�ɫ
End If
Next

For Each rng In Selection
If rng > 0 Then
rng.Interior.ColorIndex = 3    '�ı����ٿպʹ���0�ĵ�Ԫ����ɫ
End If
Next

For Each rng In Selection
If rng = 0 Then
rng.Interior.ColorIndex = 2   '��ֵ�͵���0�ĵ�Ԫ��׵�ɫ
End If

Next

End Sub






