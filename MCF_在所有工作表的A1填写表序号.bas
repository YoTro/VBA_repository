'���=������
'˵��=�����й������A1��д�����

Sub �����й������A1��д�����()
For i = 1 To Sheets.Count
Sheets(i).Cells(1, 1) = "'" & Application.WorksheetFunction.Text(0 + i, "000")
Next
End Sub




