'���=��λ����
'˵��=��λ���ݼ�����Ŀ�ֵ

Sub ��λ���ݼ�����Ŀ�ֵ()
Dim aa As Range
For Each a In ActiveSheet.UsedRange
If a Like < 0 Then
If aa Is Nothing Then
Set aa = a.Cells
Else
Set aa = Union(aa, a.Cells)
End If
End If
Next
aa.Select
End Sub






