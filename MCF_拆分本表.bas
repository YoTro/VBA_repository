'���=���˳���
'˵��=��˵��
Sub ��ֱ���() '���и��ƣ��ٶ�ƫ����ͨ���Ժ�
Dim SplitCol As String, ColNum As Integer, HeadRows As Byte
Dim arr, lastrow, i, ShtIndex
Dim only
Set only = CreateObject("scripting.dictionary") 'Set only = New Collection
'-------------
'ָ��������������С����Ը���ʵ������޸��б�
Dim tmpX
tmpX = Application.InputBox("������������������:", "ָ���������������", "E", Type:=2)
If tmpX = False Then Exit Sub
SplitCol = tmpX

'ָ�����������������򲻲�����
tmpX = Application.InputBox("ָ�����������������򲻲�����", "��������", "1", Type:=1)
If tmpX = False Then Exit Sub
HeadRows = tmpX
'-----------------
If HeadRows >= ActiveSheet.UsedRange.Rows.Count Then Exit Sub '���ָ���ı����д������������������˳�����
ColNum = Cells(1, SplitCol).Column  '���б�ת��������
lastrow = ActiveSheet.UsedRange.Rows.Count  '��ȡ��ǰ���������������
arr = Range(Cells(HeadRows + 1, SplitCol), Cells(lastrow, SplitCol)).Value  '������е����ݸ������arr
'-----------------
On Error Resume Next
For i = 1 To lastrow - HeadRows  '����arr��������
  '��ȡ���еĲ��ظ�ֵ
  If Len(arr(i, 1)) > 0 Then only.Add CStr(arr(i, 1)), CStr(arr(i, 1))
Next i
ShtIndex = ActiveSheet.Index  '��ȡ��ǰ��λ��
'-----------------
Dim ikeys
ikeys = only.keys
'-----------------
On Error Resume Next
For i = 0 To only.Count - 1
    Debug.Print Sheets(ikeys(i)).Name  '��ȡ��only������ÿ��Ԫ��ͬ���Ĺ�������������Ϊ�ж��Ƿ���ڸù�����
    If Err = 0 Then MsgBox "��ǰ�������Ѵ�����������Ŀͬ���Ĺ�����""" & ikeys(i) & """�����޷����", 64, "������ʾ": Exit Sub
    Err.Clear
Next i
'-----------------
Application.ScreenUpdating = False  '�ر���Ļ���£��ӿ�ִ���ٶ�
Application.Calculation = xlCalculationManual  '��Ϊ�ֶ����㣬�ӿ�ִ���ٶ�
For i = 0 To only.Count - 1 '������������������������only�����в��ظ�ֵ����
    Sheets.Add After:=Sheets(Sheets.Count)  '����
    Sheets(Sheets.Count).Name = ikeys(i)    '����
    Sheets(ShtIndex).Rows("1:" & HeadRows).Copy Sheets(Sheets.Count).Cells(1, 1)  '���Ʊ���
Next i
'-----------------
Sheets(ShtIndex).Select  '���ر���ֵĹ�����
For i = HeadRows + 1 To lastrow         '���и�������
  If Len(Cells(i, SplitCol)) > 0 Then  '�ų���ֵ
    With Sheets(Cells(i, SplitCol).Text).UsedRange.Rows(Sheets(Cells(i, SplitCol).Text).UsedRange.Rows.Count + 1)
          Rows(i).Copy .Cells(1)  '��һ�θ��ƣ������������ݣ���ȡ���ʽ
          .Cells = Rows(i & ":" & i).Value  '�ڶ��θ��ƣ���������ֵ
    End With
  End If
Next i   '��һ��Ϊ��ʱ������bug
'-----------------
Application.ScreenUpdating = True  '�ָ���Ļ����
Application.Calculation = xlCalculationAutomatic  '�ָ��Զ�����
MsgBox "�����ϣ�", 64, "������ʾ"
End Sub
