'���=�ϲ��Ͳ��
'˵��=ѡ�кϲ��ĵ�Ԫ��
Sub ѡ�кϲ���Ԫ��()
    On Error GoTo l_err
    
    Dim r As Range, allUsed As Range
    Dim all As Range

    Set allUsed = Intersect(ActiveSheet.UsedRange, Selection)
    
    For Each r In allUsed
        If r.MergeCells Then
        '------------------------------
            If all Is Nothing Then
                Set all = r
            Else
                Set all = Union(all, r)
            End If
        '------------------------------
        End If
    Next
    
     If Not all Is Nothing Then all.Select
    Exit Sub
l_err:
    MsgBox "��������" & Err.Description
End Sub

