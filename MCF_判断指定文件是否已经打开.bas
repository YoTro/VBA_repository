'���=
'˵��=�ж�ָ���ļ��Ƿ��Ѿ���
Sub �ж�ָ���ļ��Ƿ��Ѿ���()

    Dim i As Integer
    Dim targetFile As String
    
    targetFile = "����.xls"  '��Ҫȷ���Ƿ��Ѿ��򿪵��ļ�
    
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = targetFile Then    '�ļ�����
            MsgBox "�ļ��Ѵ�"
            Exit Sub
        End If
    Next i
    
    MsgBox "�ļ�δ��"
End Sub



