'类别=工作表
'说明=将全部表名称写到A列


Sub 将全部表名称写到A列()
k = 0
For Each Sht In Sheets
  Cells(k + 1, 1) = Sht.Name       '指定写入的行和列
  k = k + 1
Next
End Sub




