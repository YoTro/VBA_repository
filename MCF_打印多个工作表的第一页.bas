'���=��ӡ����
'˵��=��ӡ���������ĵ�һҳ

Sub ��ӡ���������ĵ�һҳ()
Dim sh As Integer
Dim x
Dim y
Dim sy
Dim syz

x = InputBox("��������ʼ����������:")
sy = InputBox("�������������������:")
y = Sheets(x).Index
syz = Sheets(sy).Index
For sh = y To syz
    Sheets(sh).Select
    Sheets(sh).PrintOut from:=1, To:=1
Next sh
End Sub





