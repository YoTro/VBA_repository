'类别=批量删除
'说明=无说明
Sub 删除选区非数字的单元格()
    On Error Resume Next
    Dim r As Range
    Set r = Intersect(ActiveSheet.UsedRange, Selection)
    
    If MsgBox("危险操作，确定删除？", vbOKCancel, "注意!") = vbCancel Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For i = r.Cells.Rows.Count To 1 Step -1
        For j = 1 To r.Cells.Columns.Count
            If (Not IsNumeric(r.Cells(i, j).Value)) Or r.Cells(i, j) = "" Then
                r.Cells(i, j).Delete xlUp
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub


