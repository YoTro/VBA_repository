'���=��λ����
'˵��=��궨λ��ָ��������A�������������һ��Ԫ

Sub ��λ��ǰ�����һ����һ��Ԫ()
    Dim colName As String
    Dim maxRow As String
    dim lastRow as String
    
    colName = Split(ActiveCell.Address, "$")(1)
    maxRow = Rows.Count

    lastRow = Range(colName & maxRow).End(xlUp).Row
    'lastRow = ActiveSheet.[a65536].End(xlUp).Row

    Range(colName  & lastRow + 1).Select
    
End Sub





