'类别=工作簿
'说明=无说明
Sub 弹出并获取文件名()  'getopenfilename
    x = Application.GetOpenFilename("all files(*.*),*.*")
    If x <> False Then
        MsgBox "你要打开的文件是:" & x
    End If
End Sub





