Sub SheetVisible()
' 一键批量取消工作表隐藏
    Dim sht As Worksheet
    '定义变量
    For Each sht In Worksheets
    '循环工作簿里的每一个工作表
        sht.Visible = xlSheetVisible
        '将工作表的状态设置为非隐藏
    Next
End Sub