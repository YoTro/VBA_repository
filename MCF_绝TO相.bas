'类别=个人常用
'说明=无说明
Sub 相TO绝() '这个是相对引用转换为绝对引用
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeFormulas)
c.Formula = Application.ConvertFormula(c.Formula, xlA1, , xlAbsolute)
Next
End Sub

Sub 绝TO相() '这个是绝对引用转换为相对引用
Dim c As Range
For Each c In Cells.SpecialCells(xlCellTypeFormulas)
c.Formula = Application.ConvertFormula(c.Formula, xlA1, , xlRelative, c)
Next
End Sub