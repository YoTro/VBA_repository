'类别=个人常用
'说明=区域录入当前日期

Sub 区域录入当前日期()
   Selection.FormulaR1C1 = Format(Now(), "yyyy-m-d")
End Sub



