'类别=
'说明=设置选区为滚动区域，使用户无法选定其他区域

Sub 设置选区为滚动区域()
    ActiveSheet.ScrollArea = Selection.Address
End Sub








