'类别=批量删除
'说明=删除选区超链接

Sub 删除选区超链接()
    If MsgBox("危险操作，确定删除？", vbOKCancel, "注意!") = vbCancel Then
        Exit Sub
    End If

    Selection.Hyperlinks.Delete

    For Each Rng In Selection
       ' Rng.Hyperlinks.Delete
    Next
End Sub






