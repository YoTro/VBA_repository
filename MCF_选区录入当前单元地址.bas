'类别=新增录入
'说明=选区录入当前单元地址




Sub 选区录入当前单元地址()
    Selection = "=ADDRESS(ROW(),COLUMN(),4,1)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub




