'���=��ֵת��
'˵��=��˵��
Sub ����Ԫ�񽻻�����()
On Error Resume Next
Dim tar As Range
Set tar = Selection
'-------------
Dim t1 As Range, t2 As Range
If tar.Areas.Count = 1 Then
    If tar.Cells.Count <> 2 Then
        MsgBox "����ѡ��2����Ԫ��"
        Exit Sub
    End If
    Set t1 = tar.Cells(1, 1)
    If tar.Rows.Count >= 2 Then '�з���
        Set t2 = tar.Cells(2, 1)
    Else
        Set t2 = tar.Cells(1, 2) '�з���
    End If
ElseIf tar.Areas.Count = 2 Then
    Set t1 = tar.Areas(1).Cells(1, 1)
    Set t2 = tar.Areas(2).Cells(1, 1)
Else
    MsgBox "����ѡ��2����Ԫ��"
    Exit Sub
End If
'-----------
Dim tmp
tmp = t1.Value
t1.Value = t2.Value
t2.Value = tmp

End Sub