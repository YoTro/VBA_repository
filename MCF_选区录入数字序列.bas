'���=����¼��
'˵��=ѡ��¼���������У���1,2,3,4...
Sub ѡ��¼����������()
    Dim i  As Integer
    'Selection.ClearContents '�������
    
    i = 0
    For Each Rng In Selection
    
        If Rng.MergeCells Then  '�Ƿ��Ǻϲ���Ԫ��
             If Rng.MergeArea.Cells.Offset.Address = Rng.Address Then  '�Ƿ��Ǻϲ���Ԫ��ĵ�һ��
                i = i + 1
                Rng.Value = i
             End If
        Else
            i = i + 1
            Rng.Value = i
        End If
    Next
    
End Sub








