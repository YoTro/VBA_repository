'类别=个人常用
'说明=无说明
Sub 提取非合并单元格整行()
    On Error GoTo l_err
    Dim r As Range
    Dim i, count
    Dim target As Range
    
    If Selection.Columns.count > 1 Then
        MsgBox "选区只允许包含一个列！"
        Exit Sub
    End If
    count = Selection.Cells.count
    
    
    Set target = Application.InputBox(prompt:="请选择单元格，用来存放剪切出来的数据(整行数据)。", Title:="结果存放", Type:=8)
    If target Is Nothing Then
        Exit Sub
    End If
    Set target = target.Cells(1, 1).EntireRow
    
    
    Application.DisplayAlerts = False
    For i = count To 1 Step -1
        Set r = Selection.Cells(i)
        If Not r.MergeCells Then
            If r.Value <> "" Then
                r.EntireRow.Cut
                target.Insert Shift:=xlDown
            End If
        End If
    Next i
    
    Application.DisplayAlerts = True
    Exit Sub
l_err:
    Application.DisplayAlerts = True
    MsgBox "发生错误：" & Err.Description
End Sub
