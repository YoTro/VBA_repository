'���=��ӡ����
'˵��=��ָ��������ִ��ȡ�����ء���ӡ�����ع�����
Sub ��ӡ���ع�����()
    On Error GoTo l_err
    Dim wsName As String
    wsName = "����1"  '��������
    
    Sheets(wsName).Visible = True
    Sheets(wsName).PrintOut Copies:=1, Collate:=True
    Sheets(wsName).Visible = False
    Exit Sub
l_err:
    msgbox "��������" & Err.Description
End Sub






