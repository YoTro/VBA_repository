'���=
'˵��=���ӡ���ָ��ʱ��ִ�к꣬���ӳټ���ִ�С���Ҫ��� ��ӵ���������ִ��

Sub ����()
    Application.OnTime ("11:45:00"), "��ʾ1"    '������
    Application.OnTime Now + TimeValue("00:00:15"), "��ʾ2"
End Sub

Sub ��ʾ1()
    msgbox "��ʾ������"
End Sub

Sub ��ʾ2()
    msgbox "��ʾ������"
End Sub





