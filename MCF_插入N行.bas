'���=���˳���
'˵��=������Ҫ��������������Ϸ����롣


Sub ����N��()
	dim count as integer
	count = Application.InputBox("����������:", "��������", 1, type:=1)
    
	If count < 1 Then Exit Sub 
        
    	Rows(ActiveCell.Row & ":" & ActiveCell.Row + count-1).Select
    	Selection.Insert Shift:=xlDown
End Sub







