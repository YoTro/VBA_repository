'���=�ϲ��Ͳ��
'˵��=��ֺϲ��ĵ�Ԫ�����

Sub ��ֵ�Ԫ�����()
    On Error GoTo l_err
    
    Dim r As Range
    Dim rt As Integer, ct As Integer
    Dim i, j
    Dim tmpV

    
    For Each r In Selection
        If r.MergeCells Then
        '------------------------------
            If r.MergeArea.Cells.Offset.Address = r.Address Then
                tmpV = r.Value
                rt = r.MergeArea.Rows.Count
                ct = r.MergeArea.Columns.Count
                '-----------------------------
                r.UnMerge
                For i = 0 To rt - 1
                    For j = 0 To ct - 1
                        r.Offset(i, j) = tmpV
                    Next j
                Next i
                
            End If
        '------------------------------
        End If
    Next
    Exit Sub
l_err:
    MsgBox "��������" & Err.Description
End Sub



