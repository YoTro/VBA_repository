'���=��ֵ��ֵ
'˵��=���ػ���ʾ��ǰ�еĿ�ֵ��
Sub ���ػ���ʾ��ǰ�еĿ�ֵ��()
    Dim colName As String
    colName = Split(ActiveCell.Address, "$")(1)

    With Range(colName & ":" & colName).SpecialCells(xlCellTypeBlanks).EntireRow
        .Hidden = Not .Hidden
    End With
End Sub




