'类别=批量删除
'说明=删除所有空白工作表
Sub 删除空白工作表()
    Dim sht As Worksheet, n As Integer, iFlag As Boolean

    If MsgBox("危险操作，确定删除？", vbOKCancel, "注意!") = vbCancel Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        'If sht.UsedRange.Cells.Count = 0 Then
        If Application.WorksheetFunction.CountA(sht.UsedRange.Cells) = 0 Then
            sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub



