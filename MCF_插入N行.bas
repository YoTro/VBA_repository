'类别=个人常用
'说明=先输入要插入的行数，在上方插入。


Sub 插入N行()
	dim count as integer
	count = Application.InputBox("请输入行数:", "输入行数", 1, type:=1)
    
	If count < 1 Then Exit Sub 
        
    	Rows(ActiveCell.Row & ":" & ActiveCell.Row + count-1).Select
    	Selection.Insert Shift:=xlDown
End Sub







