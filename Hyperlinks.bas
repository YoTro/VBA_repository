Sub Hyperlinks()
'一键生成带超链接的工作表目录

    Dim sht As Worksheet, i&, strShtName$
    Columns(1).ClearContents
   '清空A列数据
    Cells(1, 1) = "目录"
   '第一个单元格写入字符串"目录"
    i = 1
   '将i的初值设置为1.
    For Each sht In Worksheets
       '循环当前工作簿的每个工作表
        strShtName = sht.Name
        If strShtName <> ActiveSheet.Name Then
       '如果sht的名称不是当前工作表的名称则开始在当前工作表建立超链接
            i = i + 1
           '累加i
           ActiveSheet.Hyperlinks.Add anchor:=Cells(i, 1), Address:="", SubAddress:="'" & strShtName & "'!a1", TextToDisplay:=strShtName
           '建超链接
        End If
    Next
End Sub