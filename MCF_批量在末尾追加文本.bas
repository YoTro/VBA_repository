'���=���˳���
'˵��=������ĩβ�����ı�
Sub ������ĩβ׷���ı�()
    Dim r As Range
    Dim str
    str = Application.InputBox("������׷�ӵ��ı�����:", "�����ı�����")
    
    If str = False Then Exit Sub
    
    For Each r In Selection
        r = r.Value & str
    Next
End Sub







