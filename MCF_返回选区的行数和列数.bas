'类别=
'说明=返回光标选择区域的行数和列数

Sub 返回选区的行数和列数()
x = Selection.Rows.Count
y = Selection.Columns.Count
MsgBox "行数：" & x & "  列数：" & y
End Sub







