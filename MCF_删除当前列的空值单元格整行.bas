'���=��ֵ��ֵ
'˵��=ɾ����ǰ�еĿ�ֵ��Ԫ������

Sub ɾ����ǰ�еĿ�ֵ��Ԫ������()
    Dim colName As String
    colName = Split(ActiveCell.Address, "$")(1)

    With Range(colName & ":" & colName).SpecialCells(xlCellTypeBlanks).EntireRow
        .Delete
    End With
End Sub






