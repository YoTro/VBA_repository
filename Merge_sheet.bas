Attribute VB_Name = "Merge_sheet"
Sub Merge_sheet()
Application.ScreenUpdating = False
For j = 1 To Sheets.Count
If Sheets(j).Name <> ActiveSheet.Name Then
X = Range("A65536").End(xlUp).Row + 1
Sheets(j).UsedRange.Copy Cells(X, 1)
End If
Next
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "��ǰ�������µ�ȫ���������Ѿ��ϲ���ϣ�", vbInformation, "��ʾ"
End Sub

