'类别=工作簿
'说明=无说明

Option Explicit


Sub 遍历文件()
       On Error Resume Next

    Dim tar As Range
    
    Dim ik, iv
    GetFileListArr ThisWorkbook.Path, "", ik, iv  'ThisWorkbook.Path为要搜索的文件夹路径
    
    '--------------------------------------------------------
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(按列)。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Cells(1, 1).Offset(0, 0) = "文件名"
    tar.Cells(1, 1).Offset(0, 1) = "文件路径"
    
    tar.Cells(1, 1).Offset(1, 0).Resize(UBound(ik) + 1) = WorksheetFunction.Transpose(ik)
    tar.Cells(1, 1).Offset(1, 1).Resize(UBound(iv) + 1) = WorksheetFunction.Transpose(iv)
    Exit Sub

End Sub

Function GetFileListArr(folderspec As String, ext As String, ByRef ik, ByRef iv)

    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    
    GetFileList folderspec, ext, d
    
    ik = d.keys()
    iv = d.items()

End Function



Public Function GetFileList(folderspec As String, ext As String, d As Object) '遍历子文件夹的搜索
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

