'类别=个人常用
'说明=无说明
Sub 选区允许编辑()
'保护密码为空
On Error Resume Next

Dim tar As Range
Set tar = Selection
tar.Worksheet.Unprotect '取消保护工作表

If tar.Worksheet.Cells.Locked = True Then
    tar.Worksheet.Cells.Locked = False
End If
tar.Locked = False
tar.FormulaHidden = False
tar.Worksheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
