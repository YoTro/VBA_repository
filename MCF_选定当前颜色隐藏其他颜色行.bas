'���=��λ����
'˵��=ѡ����ǰ��ɫ����������ɫ����
Sub ѡ����ǰ��ɫ����������ɫ��() '˼·����:��������ɫ֮��ȫ������
    Dim UseRow, AC, i '����ѡ��һ������ɫ֮��Ԫ��Ȼ���к꣬������ɫ����������
    UseRow = Cells.SpecialCells(xlCellTypeLastCell).Row 'SpecialCells(xlCellTypeLastCell)��ʾ�����������һ����Ԫ��
    If ActiveCell.Row > UseRow Then
        MsgBox "����Ҫɸѡ������ѡ��һ������ɫ֮��Ԫ��", vbExclamation, "����"
    Else
        AC = ActiveCell.Column
        Cells.EntireRow.Hidden = False '��ʾ������
        
        For i = 2 To UseRow
            If Cells(i, AC).Interior.ColorIndex <> ActiveCell.Interior.ColorIndex Then
                Cells(i, AC).EntireRow.Hidden = True '���2��������֮��Ԫ�������֮��ɫ�����ڵ�ǰ��Ԫ����ɫ����������
            End If
        Next
    End If
End Sub



