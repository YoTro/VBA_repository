'���=����ɾ��
'˵��=ɾ��ѡ��������

Sub ɾ��ѡ��������()
    If MsgBox("Σ�ղ�����ȷ��ɾ����", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    Selection.Hyperlinks.Delete

    For Each Rng In Selection
       ' Rng.Hyperlinks.Delete
    Next
End Sub






