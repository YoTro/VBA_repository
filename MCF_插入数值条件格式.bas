'类别=定位引用
'说明=插入数值条件格式

Sub 插入数值条件格式()

    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:=">70"
    Selection.FormatConditions(1).Interior.ColorIndex = 45
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="55"
    Selection.FormatConditions(2).Interior.ColorIndex = 39
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="60"
    Selection.FormatConditions(3).Interior.ColorIndex = 34
End Sub






