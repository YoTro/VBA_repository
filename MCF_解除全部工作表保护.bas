'类别=工作表
'说明=解除全部工作表保护

Sub 解除全部工作表保护()
    Dim n As Integer
    For n = 1 To Sheets.Count
        Sheets(n).Unprotect
    Next n
End Sub


