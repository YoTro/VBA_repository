Sub NewSht()
    '根据A列数据批量建立工作表的代码如下
    Dim Sht As Worksheet, Rng As Range
    Dim Sn, t$
    Set Rng = Range("a2:a" & Cells(Rows.Count, 1).End(xlUp).Row)
    '将工作表名称所在的单元格区域赋值给变量Rng，单元格A1是标题，不读入
    On Error Resume Next
    '当代码出错时继续运行
    For Each Sn In Rng
    '遍历Rng(工作表名称集合）
        t = Sn
        '还记得这里我们为什么用这句代码吗？
        Set Sht = Sheets(t)
        '当工作簿不存在工作表Sheets(t)时，这句代码会出错，然后……
        If Err Then
        '如果代码出错，说明不存在工作表Sheets(t)，则新建工作表
            Worksheets.Add , Sheets(Sheets.Count)
            '新建一个工作表，位置放在所有已存在工作表的后面
            ActiveSheet.Name = t
            '新建的工作表必然是活动工作表，为之命名
            Err.Clear
            '清除错误状态
        End If
    Next
    Rng.Parent.Activate
    '重新激活名称数据所在的工作表
End Sub