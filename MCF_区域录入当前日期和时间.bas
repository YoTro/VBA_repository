'类别=新增录入
'说明=区域录入当前日期和时间

Sub 区域录入当前日期和时间()
    Selection.FormulaR1C1 = Format(Now(), "yyyy-m-d h:mm:ss")
End Sub



