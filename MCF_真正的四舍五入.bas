'类别=数值转换
'说明=无说明
Sub 真正的四舍五入()
    Dim r As Range
    Dim str
    
    Dim bitnum As Double
    Dim tmp As Double
    '-----------------------------
    str = Application.InputBox("请输入要保留的小数位数", "输入", "2")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    
    bitnum = CDbl(str)
    If bitnum < 0 Then Exit Sub
    '-----------------------------
    For Each r In Selection
        If IsNumeric(r.Value) Then
            tmp = Application.WorksheetFunction.Round(r.Value, bitnum)
            r.Value = tmp
        End If
    Next

End Sub
