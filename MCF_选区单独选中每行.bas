'���=��λ����
'˵��=��˵��
Sub ѡ������ѡ��ÿ��()
Dim tar As Range, r As Range
Set tar = Selection
Dim rlt As String
rlt = ""

Set tar = tar.Cells.SpecialCells(xlCellTypeVisible)
If tar.Rows.Count >= Rows.Count Then
    MsgBox "̫�����ˣ�"
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
    MsgBox "ѡ��̫�����ˡ�"
    Exit Sub
End If
If rlt <> "" Then Range(rlt).Select

End Sub

