'类别=定位引用
'说明=返回当前列数据的最大行数

Sub 返回当前列数据的最大行数()
    n =Cells( Rows.Count,ActiveCell.Column).End(xlUp).Row
    msgbox n
End Sub






