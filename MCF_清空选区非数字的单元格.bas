'���=����ɾ��
'˵��=��շ����ֵĵ�Ԫ��
Sub ���ѡ�������ֵĵ�Ԫ��()
    Dim r As Range
    If MsgBox("Σ�ղ�����ȷ����գ�", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    For Each r In Selection
        If Not IsNumeric(r.Value) Then
            r = ""
        End If
    Next
End Sub









