'类别=工作簿
'说明=无说明
Sub 录入本工作簿路径()
	On Error Resume Next
	ActiveCell.Value = ThisWorkbook.FullName

End Sub
