'���=������
'˵��=��˵��
Sub ��������ȡ�ļ���()  'getopenfilename
    x = Application.GetOpenFilename("all files(*.*),*.*")
    If x <> False Then
        MsgBox "��Ҫ�򿪵��ļ���:" & x
    End If
End Sub





