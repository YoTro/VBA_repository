'���=������
'˵��=�ϲ�������������

Sub �ϲ�������������()
sp = InputBox("��������֮�䣬������У�������Ĭ��Ϊ0")
If sp = "" Then
  sp = 0
End If

st = InputBox("����ӵڼ��п�ʼ�ϲ���������Ĭ��Ϊ2")
If st = "" Then
   st = 2
End If

Sheets(1).Select
Sheets.Add
  
  If st > 1 Then
    Sheets(2).Select
    Rows("1:" & CStr(st - 1)).Select
    Selection.Copy
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Paste
  y = st - 1
  End If
  
For i = 2 To Sheets.Count
    Sheets(i).visible = true
  Sheets(i).Select
     For v = 1 To 256
        zd = Cells(65535, v).End(xlUp).Row
        If zd > x Then
           x = zd
        End If
     Next v

  If y + x - st + 1 + sp > 65536 Then
  MsgBox "����̫�࣬���ϲ�ǰ" & i - 2 & "��������ݣ�����������Ƶ��¹����������ô˳���ϲ���"
  Else:
  
  Rows(st & ":" & x).Select
  Selection.Copy
  Sheets(1).Select
  Range("A" & CStr(y + 1)).Select
  ActiveSheet.Paste
  
  Sheets(i).Select
  Range("A1").Select                        'ȡ����Ԫ��ȫѡ״̬��
  Application.CutCopyMode = False           '�������Ƶ����ݡ�
  End If
  
  y = y + x - st + 1 + sp
  x = 0
Next i

Sheets(1).Select
Range("A1").Select                          '�������A1��
MsgBox "����Ǻϲ���ı���������"

End Sub







