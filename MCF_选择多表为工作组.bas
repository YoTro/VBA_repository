'���=������
'˵��=ѡ����Ϊ������

Sub ѡ����Ϊ������()
Dim Wks As Worksheet, shtCnt As Integer
Dim arr() As Variant, i As Integer, m As Integer, m1 As Integer, m2 As Integer
shtCnt = ThisWorkbook.Sheets.Count 'ȡ�ù���������
ReDim arr(1 To shtCnt) 'Ԥ��������
i = 0
m = 1  'ѭ���Ĵ���
m1 = 0 '�ҵ����ѭ���Ĵ���
m2 = 0 '�ҵ��յ�ѭ���Ĵ���
For Each Wks In ThisWorkbook.Sheets '�����й�������ѭ��
    If Wks.Name = "A2" Then   '�������е�һ������������
        i = i + 1
        arr(i) = Wks.Name '�����������ƴ������
        m1 = m
    End If
    If Wks.Name Like "A7" Then    '�����������һ��������������
        i = i + 1
        arr(i) = Wks.Name '�����������ƴ������
        m2 = m
        Exit For
    End If
    If i > 0 And m > m1 Then
        i = i + 1
        arr(i) = Wks.Name '�����������ƴ������
    End If
    m = m + 1
Next
If m2 > m1 Then '������ڷ��������Ĺ���������
    ReDim Preserve arr(1 To i) '�ض�������
    ThisWorkbook.Sheets(arr).Select 'ѡ�з������������й�����
End If
End Sub



