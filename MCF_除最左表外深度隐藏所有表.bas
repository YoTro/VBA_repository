'类别=工作表
'说明=除最左边工作表外深度隐藏所有表

Sub 除最左表外深度隐藏所有表()
    For i = 2 To ThisWorkbook.Sheets.Count
        Sheets(i).Visible = xlSheetVeryHidden
    Next
End Sub


