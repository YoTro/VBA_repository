'类别=打印工具
'说明=打印多个工作表的第一页

Sub 打印多个工作表的第一页()
Dim sh As Integer
Dim x
Dim y
Dim sy
Dim syz

x = InputBox("请输入起始工作表名字:")
sy = InputBox("请输入结束工作表名字:")
y = Sheets(x).Index
syz = Sheets(sy).Index
For sh = y To syz
    Sheets(sh).Select
    Sheets(sh).PrintOut from:=1, To:=1
Next sh
End Sub





