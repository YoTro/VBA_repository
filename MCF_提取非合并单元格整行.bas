'���=���˳���
'˵��=��˵��
Sub ��ȡ�Ǻϲ���Ԫ������()
    On Error GoTo l_err
    Dim r As Range
    Dim i, count
    Dim target As Range
    
    If Selection.Columns.count > 1 Then
        MsgBox "ѡ��ֻ�������һ���У�"
        Exit Sub
    End If
    count = Selection.Cells.count
    
    
    Set target = Application.InputBox(prompt:="��ѡ��Ԫ��������ż��г���������(��������)��", Title:="������", Type:=8)
    If target Is Nothing Then
        Exit Sub
    End If
    Set target = target.Cells(1, 1).EntireRow
    
    
    Application.DisplayAlerts = False
    For i = count To 1 Step -1
        Set r = Selection.Cells(i)
        If Not r.MergeCells Then
            If r.Value <> "" Then
                r.EntireRow.Cut
                target.Insert Shift:=xlDown
            End If
        End If
    Next i
    
    Application.DisplayAlerts = True
    Exit Sub
l_err:
    Application.DisplayAlerts = True
    MsgBox "��������" & Err.Description
End Sub
