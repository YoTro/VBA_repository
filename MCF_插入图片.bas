'���=����ͼƬ
'˵��=��˵��
Sub ����ͼƬ()  'addpicture

    Dim fpath
    fpath = Application.GetOpenFilename("all files(*.*),*.*")
    If fpath = False Then
        Exit Sub
    End If

    ActiveSheet.Shapes.AddPicture fpath, True, True, 30, 10, 300, 225

End Sub

