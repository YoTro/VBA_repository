'类别=定位引用
'说明=选区半角内容变红
Sub 选区半角内容变红()
On Error Resume Next
Dim rg As Range, i As Long

Application.ScreenUpdating = False

For Each rg In Selection.SpecialCells(xlCellTypeConstants, 3)
	For i = 1 To Len(rg)
		If Asc(Mid(rg, i, 1)) > 0 Then rg.Characters(i).Font.ColorIndex = 3
	Next
Next

Application.ScreenUpdating = True
End Sub





