'类别=个人常用
'说明=无说明
Sub 批量删除条件格式()

    If MsgBox("危险操作，确定删除本工作簿的所有条件格式？", vbOKCancel, "注意!") = vbCancel Then
        Exit Sub
    End If

    Dim sh As Worksheet
    For Each sh In Worksheets
        sh.Cells.FormatConditions.Delete
    Next
    msgbox "完成"
End Sub
