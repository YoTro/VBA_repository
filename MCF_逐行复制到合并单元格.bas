'���=���˳���
'˵��=��˵��

Sub ���и��Ƶ��ϲ���Ԫ��()
Dim tar As Range, org As Range
Set org = Selection
If org.Cells.Count >= 65536 Then
    MsgBox "ѡ��ĵ�Ԫ��̫����(����65536��)"
    Exit Sub
End If


Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��", Title:="������", Type:=8)
If tar Is Nothing Then
    Exit Sub
End If
Set tar = tar.Cells(1, 1)

If org.Cells.Count = 1 Then
    tar.Cells(1, 1) = org.Cells(1, 1).Value
    Exit Sub
End If

Dim r As Range
Dim data, i, j
data = org
For i = 1 To UBound(data, 1) '����
    Set r = tar.Offset(i - 1, 0)
    For j = 1 To UBound(data, 2)
        Set r = GetRightUnMergeRange(r)
        If Not r Is Nothing Then
            r.Value = data(i, j)
            Set r = r.Offset(0, 1)
        End If
    Next
Next i





End Sub

Function GetRightUnMergeRange(tar As Range) As Range
On Error Resume Next
Dim r As Range

For i = 0 To Rows.Count
    Set r = tar.Offset(0, i)
    
    If r.MergeCells Then '�ϲ�
        If r.MergeArea.Cells.Offset.Address = r.Address Then  '�׸�
            Set GetRightUnMergeRange = r
            Exit Function
        End If
    Else '�Ǻϲ�
        Set GetRightUnMergeRange = r
        Exit Function
    End If
Next

Set GetRightUnMergeRange = Nothing
End Function