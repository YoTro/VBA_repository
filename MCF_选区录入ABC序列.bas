'���=����¼��
'˵��=ѡ��¼��ABC����


Sub ѡ��¼��ABC����()
    Dim i  As Integer
    'Selection.ClearContents '�������
    
    i = 0
    For Each Rng In Selection
        If Rng.MergeCells Then  '�Ƿ��Ǻϲ���Ԫ��
             If Rng.MergeArea.Cells.Offset.Address = Rng.Address Then  '�Ƿ��Ǻϲ���Ԫ��ĵ�һ��
                Rng.Value = Chr(65 + i)
                i = (i + 1) Mod 26
             End If
        Else
            Rng.Value = Chr(65 + i)
            i = (i + 1) Mod 26
        End If
    Next
    
End Sub







