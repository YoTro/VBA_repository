'���=������
'˵��=��˵��
Sub ���Ʊ���N��()
    On Error Resume Next
    Dim str, prefix As String
    
    Dim num As Integer
    Dim ws As Worksheet, wb As Workbook
    Set ws = ActiveSheet
    Set wb = ws.Parent
    '-----------------------------
    str = Application.InputBox("������Ҫ���ƶ��ٷ�", "����", "5")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    prefix = Application.InputBox("���빤��������ǰ׺", "����", "Sheet")
    
    num = CInt(str)
    If num < 0 Then Exit Sub
    '-----------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 1 To num
        ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)
        wb.Worksheets(wb.Worksheets.Count).Name = prefix & i
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "���"

End Sub