'类别=工作表
'说明=在所有工作表的A1填写表序号

Sub 在所有工作表的A1填写表序号()
For i = 1 To Sheets.Count
Sheets(i).Cells(1, 1) = "'" & Application.WorksheetFunction.Text(0 + i, "000")
Next
End Sub




