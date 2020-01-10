'类别=打印工具
'说明=查找A列文本循环插入分页符

Sub 循环插入分页符()
    'Selection = Workbooks("临时表").Sheets("表2").Range("A1") 调用指定地址内容
  
    Dim i As Long
    Dim times As Long
    times = Application.WorksheetFunction.CountIf(Sheet1.Range("a:a"), "分页")
    'times代表循环次数，执行前把times赋值即可(不可小于1，不可大于2147483647)
    For i = 1 To times
	Call 插入分页符
    Next i
End Sub


Sub 插入分页符()
    Cells.Find(What:="分页", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False) _
        .Activate
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
End Sub


Sub 取消原分页()
    Cells.Select
    ActiveSheet.ResetAllPageBreaks
End Sub





