'类别=
'说明=判断指定文件是否已经打开
Sub 判断指定文件是否已经打开()

    Dim i As Integer
    Dim targetFile As String
    
    targetFile = "函数.xls"  '你要确定是否已经打开的文件
    
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = targetFile Then    '文件名称
            MsgBox "文件已打开"
            Exit Sub
        End If
    Next i
    
    MsgBox "文件未打开"
End Sub



