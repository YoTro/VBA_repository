'���=��ֵת��
'˵��=��˵��
Sub �޸��������Ļ��з�()
Dim r As Range
Dim s As String

For Each r In Selection.Cells
    s = r.Value
    's = Replace(s, vbCr, vbCrLf)
    s = Replace(s, vbLf, vbCrLf)
    r.Value = s
Next
msgbox "���"
End Sub
