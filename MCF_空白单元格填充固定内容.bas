'���=����¼��
'˵��=��˵��

Sub �հ׵�Ԫ�����̶�����()
    Dim str
    Dim r As Range
    Dim cols, rows
    Dim i, j
    
    str = Application.InputBox("�������ı�����:", "�����ı�����")
    
    If str = False Then Exit Sub
    

    If Selection.Cells.Count < 1 Then
        MsgBox "��ѡ��һ������"
        Exit Sub
    End If
    
    If Selection.Areas.Count > 1 Then Exit Sub
    rows = Selection.Cells.rows.Count
    cols = Selection.Cells.Columns.Count
    
    For i = 1 To rows
        For j = 1 To cols
            Set r = Selection.Cells(i, j)
            If r.Value = "" Then
                r = str
            End If
        Next j
    Next i
End Sub


