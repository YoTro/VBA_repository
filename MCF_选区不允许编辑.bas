'���=���˳���
'˵��=��˵��
Sub ѡ��������༭()
'��������Ϊ��
On Error Resume Next
Dim tar As Range
Set tar = Selection
tar.Worksheet.Unprotect 'ȡ������������
If tar.Worksheet.Cells.Locked = True Then
    tar.Worksheet.Cells.Locked = False
End If

tar.Locked = True
tar.FormulaHidden = False
tar.Worksheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub