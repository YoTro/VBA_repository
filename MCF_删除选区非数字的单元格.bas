'���=����ɾ��
'˵��=��˵��
Sub ɾ��ѡ�������ֵĵ�Ԫ��()
    On Error Resume Next
    Dim r As Range
    Set r = Intersect(ActiveSheet.UsedRange, Selection)
    
    If MsgBox("Σ�ղ�����ȷ��ɾ����", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For i = r.Cells.Rows.Count To 1 Step -1
        For j = 1 To r.Cells.Columns.Count
            If (Not IsNumeric(r.Cells(i, j).Value)) Or r.Cells(i, j) = "" Then
                r.Cells(i, j).Delete xlUp
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub


