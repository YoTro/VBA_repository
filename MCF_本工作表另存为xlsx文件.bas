'���=������
'˵��=��������������ļ���Excel��ǰĬ��Ŀ¼.
Sub �����������Ϊxlsx�ļ�()
    Dim curWs As Worksheet
    Dim wb As Workbook

    Set curWs = ActiveSheet
    Set wb = Workbooks.Add
    
    curWs.Copy before:=wb.Worksheets(1)
    'wb.Worksheets(1).Name = curWs.Name
    
    wb.SaveAs ThisWorkbook.Path & "\" & curWs.Name & ".xlsx"
    wb.Close
End Sub




