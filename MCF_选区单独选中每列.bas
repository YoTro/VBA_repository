'���=��λ����
'˵��=��˵��

Sub ѡ������ѡ��ÿ��()
Dim tar As Range, r As Range
Set tar = Selection
Dim rlt As String
rlt = ""

Set tar = tar.Cells.SpecialCells(xlCellTypeVisible)
If tar.Columns.Count >= Columns.Count Then
    MsgBox "̫�����ˣ�"
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
    MsgBox "ѡ��̫�����ˡ�"
    Exit Sub
End If
If rlt <> "" Then Range(rlt).Select

End Sub
