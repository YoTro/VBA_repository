'类别=新增录入
'说明=当前单元录入计算机用户名

Sub 当前单元录入计算机用户名()
   Selection = Environ("Username")
  'Selection = Workbooks("临时表").Sheets("表2").Range("A1") 调用指定地址内容
End Sub




