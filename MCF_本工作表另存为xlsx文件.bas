'类别=工作簿
'说明=本工作表单独另存文件到Excel当前默认目录.
Sub 本工作表另存为xlsx文件()
    Dim curWs As Worksheet
    Dim wb As Workbook

    Set curWs = ActiveSheet
    Set wb = Workbooks.Add
    
    curWs.Copy before:=wb.Worksheets(1)
    'wb.Worksheets(1).Name = curWs.Name
    
    wb.SaveAs ThisWorkbook.Path & "\" & curWs.Name & ".xlsx"
    wb.Close
End Sub




