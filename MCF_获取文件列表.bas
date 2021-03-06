'类别=
'说明=无说明
Option Explicit


Public Sub 获取文件列表()
    Dim ik, iv
    GetFileListArr ThisWorkbook.Path & "\新建文件夹", "", ik, iv
    
    'MsgBox iv(0)

End Sub

Public Function GetFileListArr(folderspec As String, ext As String, ByRef ik, ByRef iv)

    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    
    GetFileList folderspec, ext, d
    
    ik = d.keys() '从0开始
    iv = d.items()

End Function



Private Function GetFileList(folderspec As String, ext As String, d As Object) '遍历子文件夹的搜索
    On Error GoTo l_err
    
    Dim fs
    
    Dim fd As Object
    Dim t As Object
    Dim SerchFD, SerchFF

    
    Set fs = CreateObject("Scripting.FileSystemObject")

    Dim extLen As Integer
    extLen = Len(ext)
    
    Set fd = fs.GetFolder(folderspec)
    Set SerchFD = fd.SubFolders   '定义文件夹对象
    Set SerchFF = fd.Files       '定义文件对象
    '遍历子文件夹
    For Each t In SerchFD
        '加上下面这句可遍历子文件夹
        Call GetFileList(folderspec & "\" & t.Name, ext, d)
    Next
    
    '遍历文件
    For Each t In SerchFF
        If Right(t.Name, extLen) = ext Then
            d.Add t.Name, folderspec & "\" & t.Name
        End If
    Next

    Exit Function
l_err:
End Function
