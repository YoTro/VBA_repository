'���=����¼��
'˵��=

Sub �հ׵�Ԫ���������߷ǿ�ֵ()
    Dim r As Range, tmp
    Dim cols, rows
    Dim i, j
    
    If Selection.Cells.Count <= 1 Then
        MsgBox "��ѡ��һ������"
        Exit Sub
    End If
    
    If Selection.Areas.Count > 1 Then Exit Sub
    rows = Selection.Cells.rows.Count
    cols = Selection.Cells.Columns.Count
    
    For i = 1 To rows
        tmp = ""
        For j = 1 To cols
            Set r = Selection.Cells(i, j)
            If r.Value = "" Then
                r = tmp
            Else
                tmp = r
            End If
        Next j
    Next i
End Sub
