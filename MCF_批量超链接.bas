'���=���˳���
'˵��=��˵��
Sub ����������()
    On Error Resume Next
    Dim r1 As Range, r2 As Range, tar As Range
    Set r1 = Application.InputBox(prompt:="��ѡ���ļ�·����������", Title:="�ļ�·��", Type:=8)
    If r1 Is Nothing Then
        Exit Sub
    End If
    '----------------
    Set r2 = Application.InputBox(prompt:="��ѡ������Ҫ����ʾ�ı���������", Title:="��ʾ�ı�", Type:=8)
    If r2 Is Nothing Then
        Exit Sub
    End If
    '----------------
    If r1.Rows.count = r2.Rows.count And r1.Columns.count = r2.Columns.count Then
    Else
        MsgBox "������Ĵ�С��һ��"
        Return
    End If
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(һ������)��", Title:="������", Type:=8)
    If tar Is Nothing Then
        Exit Sub
    End If
    tar = tar.Resize(r1.Rows.count, r1.Columns.count)
    '----------------
    Dim i, j
    Dim txt1 As String, txt2 As String
    For i = 1 To r1.Rows.count
        For j = 1 To r1.Columns.count
            txt1 = r1.Cells(i, j).Value
            txt2 = r2.Cells(i, j).Value
            ActiveSheet.Hyperlinks.Add Anchor:=tar.Cells(i, j), Address:=txt1, TextToDisplay:=txt2
        Next
    Next
    '----------------
  
   MsgBox "���"
End Sub