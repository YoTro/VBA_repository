'类别=工作表
'说明=无说明
Sub 复制本表N份()
    On Error Resume Next
    Dim str, prefix As String
    
    Dim num As Integer
    Dim ws As Worksheet, wb As Workbook
    Set ws = ActiveSheet
    Set wb = ws.Parent
    '-----------------------------
    str = Application.InputBox("请输入要复制多少份", "输入", "5")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    prefix = Application.InputBox("输入工作表名的前缀", "输入", "Sheet")
    
    num = CInt(str)
    If num < 0 Then Exit Sub
    '-----------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 1 To num
        ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)
        wb.Worksheets(wb.Worksheets.Count).Name = prefix & i
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "完成"

End Sub