'类别=工作表
'说明=撤消工作表保护并取消密码

Sub 撤消工作表保护并取消密码()
    ActiveSheet.Unprotect Password:=123456
End Sub


