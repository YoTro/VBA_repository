'���=���˳���
'˵��=��Ԫ�����÷�ʽ�ο�
Sub ��Ԫ�����÷�ʽ()
    '------------------------------
    [A1].Select
    [A1:b3].Select
    
    
    Range("A1", "B3").Select
    Range("A1:B3").Select
    Range(Range("A1"), [B3]).Select
    Range("A1", [B3]).Select
    
    Range("A1").Offset(1, 1).Select
    Range("A2:B3").Cells(1, 1).Select
    
    Cells(1, 1).Select
    
    Range("A1", Range("A" & Rows.Count).End(xlUp)).Select
    Range("A1", Cells(1, Columns.Count).End(xlToLeft)).Select
    '------------------------------
    Rows(3).Select
    Rows("3").Select
    Rows("3:3").Select
    
    Columns(2).Select
    Columns("B").Select
    Columns("B:B").Select
    '------------------------------
End Sub

