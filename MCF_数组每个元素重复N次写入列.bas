'类别=个人常用
'说明=数组每个元素重复N次写入列
Sub 数组每个元素重复N次写入列()
   
    Dim i, j, pos As Integer
    Dim strs, count, arr
    

    strs = Application.InputBox(prompt:="输入数组用逗号隔开，如 AA,BB,CC):", Type:=2)
    count = Application.InputBox(prompt:="输入要重复的次数:", Type:=1)
    arr = Split(strs, ",")
    
    
    pos = 0
    For i = 0 To UBound(arr)
        For j = 1 To count
            ActiveCell.Offset(pos, 0) = arr(i)
            pos = pos + 1
        Next j
    Next i
End Sub





