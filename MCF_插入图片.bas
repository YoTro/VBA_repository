'类别=对象图片
'说明=无说明
Sub 插入图片()  'addpicture

    Dim fpath
    fpath = Application.GetOpenFilename("all files(*.*),*.*")
    If fpath = False Then
        Exit Sub
    End If

    ActiveSheet.Shapes.AddPicture fpath, True, True, 30, 10, 300, 225

End Sub

