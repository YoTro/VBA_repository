'���=��ӡ����
'˵��=����A���ı�ѭ�������ҳ��

Sub ѭ�������ҳ��()
    'Selection = Workbooks("��ʱ��").Sheets("��2").Range("A1") ����ָ����ַ����
  
    Dim i As Long
    Dim times As Long
    times = Application.WorksheetFunction.CountIf(Sheet1.Range("a:a"), "��ҳ")
    'times����ѭ��������ִ��ǰ��times��ֵ����(����С��1�����ɴ���2147483647)
    For i = 1 To times
	Call �����ҳ��
    Next i
End Sub


Sub �����ҳ��()
    Cells.Find(What:="��ҳ", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False) _
        .Activate
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
End Sub


Sub ȡ��ԭ��ҳ()
    Cells.Select
    ActiveSheet.ResetAllPageBreaks
End Sub





