'���=�ظ�ֵ�����ֵ
'˵��=ѡ�����ݴ���˳��


Sub ѡ�����ݴ���˳��()
    Dim ar, i, ii
    Dim tmp, tr, tc
    
    If Selection.Areas.count > 1 Then Exit Sub
    If Selection.Cells.count > Columns.count Then
        MsgBox "��ѡ����������"
        Exit Sub
    End If
    
    ar = Selection
    
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tr = Int(Rnd * UBound(ar) + 1)
            tc = Int(Rnd * UBound(ar, 2) + 1)
            
            tmp = ar(tr, tc)
            ar(tr, tc) = ar(i, ii)
            ar(i, ii) = tmp
        Next
    Next
    

    Selection = ar
End Sub








