'���=���˳���
'˵��=��˵��
Sub ����ɾ��������ʽ()

    If MsgBox("Σ�ղ�����ȷ��ɾ����������������������ʽ��", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    Dim sh As Worksheet
    For Each sh In Worksheets
        sh.Cells.FormatConditions.Delete
    Next
    msgbox "���"
End Sub
