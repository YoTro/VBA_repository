'类别=新增录入
'说明=当前单元录入计算机名

Sub 当前单元录入计算机名()
   Selection = Environ("COMPUTERNAME")
  'Selection = Workbooks("临时表").Sheets("表2").Range("A1") 调用指定地址内容
End Sub




