'类别=新增录入
'说明=区域录入对应单元地址

Sub 区域录入对应单元地址()
    For Each mycell In Selection
        mycell.FormulaR1C1 = mycell.Address
    Next
End Sub




