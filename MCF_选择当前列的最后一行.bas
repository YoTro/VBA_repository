'���=��λ����
'˵��=ѡ��ǰ�е����һ��
Sub ѡ��ǰ�е����һ��()
    Dim colName As String
    Dim maxRow As String
    
    colName = Split(ActiveCell.Address, "$")(1)
    maxRow = Rows.Count
    Range(colName & maxRow).End(xlUp).EntireRow.Select
    
    
    'ActiveCell.End(xlDown).EntireRow.Select
    'ActiveCell.SpecialCells(xlCellTypeLastCell).EntireRow.Select
End Sub



