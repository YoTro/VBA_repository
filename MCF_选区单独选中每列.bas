'类别=定位引用
'说明=无说明

Sub 选区单独选中每列()
Dim tar As Range, r As Range
Set tar = Selection
Dim rlt As String
rlt = ""

Set tar = tar.Cells.SpecialCells(xlCellTypeVisible)
If tar.Columns.Count >= Columns.Count Then
    MsgBox "太多列了！"
    Exit Sub
End If

For Each r In tar.Columns
    If rlt = "" Then
        rlt = r.Address
    Else
        rlt = rlt & "," & r.Address
    End If
    
Next

If Len(rlt) >= 255 Then
    MsgBox "选中太多列了。"
    Exit Sub
End If
If rlt <> "" Then Range(rlt).Select

End Sub
