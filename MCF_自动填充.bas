'类别=
'说明=无说明
Sub 自动填充()
    Range("A1:A2").AutoFill Destination:=Range("A1:A12")
    '在A1：A2中填入1，2或一月，二月等
End Sub

