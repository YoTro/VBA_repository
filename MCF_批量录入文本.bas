'���=���˳���
'˵��=����¼���ı�
Sub ����¼���ı�()
    Dim r As Range
    Dim str
    str = Application.InputBox("�������ı�����:", "�����ı�����")
    
    If str = False Then Exit Sub
    
    For Each r In Selection
        r =  str
    Next
End Sub





