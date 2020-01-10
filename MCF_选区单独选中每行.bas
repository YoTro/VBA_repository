'类别=定位引用
'说明=无说明
Sub 选区单独选中每行()
Dim tar As Range, r As Range
Set tar = Selection
Dim rlt As String
rlt = ""

Set tar = tar.Cells.SpecialCells(xlCellTypeVisible)
If tar.Rows.Count >= Rows.Count Then
    MsgBox "太多行了！"
    Exit Sub
End If

For Each r In tar.Rows
    If rlt = "" Then
        rlt = r.Address(False, False, xlA1)
    Else
        rlt = rlt & "," & r.Address(False, False, xlA1)
    End If
    
Next

If Len(rlt) >= 255 Then
    MsgBox "选中太多行了。"
    Exit Sub
End If
If rlt <> "" Then Range(rlt).Select

End Sub

