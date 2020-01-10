'类别=合并和拆分
'说明=合并单元格并连接字符串
Sub 合并单元格并连接字符串()

On Error GoTo l_err
Dim Strtotal
Dim r As Range

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each r In Selection
    Strtotal = Strtotal & r.Value
Next

Selection.Merge

With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "'" & Strtotal  '在合并数据前加 '号
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Exit Sub

l_err:
    MsgBox "Err: " & Err.Description

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub



