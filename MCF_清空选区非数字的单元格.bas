'类别=批量删除
'说明=清空非数字的单元格
Sub 清空选区非数字的单元格()
    Dim r As Range
    If MsgBox("危险操作，确定清空？", vbOKCancel, "注意!") = vbCancel Then
        Exit Sub
    End If

    For Each r In Selection
        If Not IsNumeric(r.Value) Then
            r = ""
        End If
    Next
End Sub









