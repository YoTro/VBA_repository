'类别=工作表
'说明=选择多表为工作组

Sub 选择多表为工作组()
Dim Wks As Worksheet, shtCnt As Integer
Dim arr() As Variant, i As Integer, m As Integer, m1 As Integer, m2 As Integer
shtCnt = ThisWorkbook.Sheets.Count '取得工作表总数
ReDim arr(1 To shtCnt) '预定义数组
i = 0
m = 1  '循环的次数
m1 = 0 '找到起点循环的次数
m2 = 0 '找到终点循环的次数
For Each Wks In ThisWorkbook.Sheets '在所有工作表中循环
    If Wks.Name = "A2" Then   '工作组中第一个工作表名称
        i = i + 1
        arr(i) = Wks.Name '将工作表名称存进数组
        m1 = m
    End If
    If Wks.Name Like "A7" Then    '工作组中最后一个个工作表名称
        i = i + 1
        arr(i) = Wks.Name '将工作表名称存进数组
        m2 = m
        Exit For
    End If
    If i > 0 And m > m1 Then
        i = i + 1
        arr(i) = Wks.Name '将工作表名称存进数组
    End If
    m = m + 1
Next
If m2 > m1 Then '如果存在符合条件的工作表名称
    ReDim Preserve arr(1 To i) '重定义数组
    ThisWorkbook.Sheets(arr).Select '选中符合条件的所有工作表
End If
End Sub



