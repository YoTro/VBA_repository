'���=���˳���
'˵��=��˵��
Sub ��TO��() '������������ת��Ϊ��������
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeFormulas)
c.Formula = Application.ConvertFormula(c.Formula, xlA1, , xlAbsolute)
Next
End Sub

Sub ��TO��() '����Ǿ�������ת��Ϊ�������
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeFormulas)
c.Formula = Application.ConvertFormula(c.Formula, xlA1, , xlRelative, c)
Next
End Sub