'类别=新增录入
'说明=区域录入当前数字日期

Sub 区域录入当前数字日期()
   Selection.FormulaR1C1 = Format(Now(), "yyyymmdd")
End Sub



