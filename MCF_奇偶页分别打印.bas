'���=��ӡ����
'˵��=��żҳ�ֱ��ӡ

Sub ��żҳ�ֱ��ӡ()
  Dim i%, Ps%
  Ps = ExecuteExcel4Macro("GET.DOCUMENT(50)") '��ҳ��
  MsgBox "���ڴ�ӡ����ҳ,��ȷ����ʼ."
  For i = 1 To Ps Step 2
    ActiveSheet.PrintOut from:=i, To:=i
  Next i
  MsgBox "���ڴ�ӡż��ҳ,��ȷ����ʼ."
  For i = 2 To Ps Step 2
    ActiveSheet.PrintOut from:=i, To:=i
  Next i
End Sub




