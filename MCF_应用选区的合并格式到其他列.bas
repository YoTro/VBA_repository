'���=�ϲ��Ͳ��
'˵��=��ѡ���еĺϲ���ʽ�ϲ�������
Sub Ӧ��ѡ���ĺϲ���ʽ��������()
    On Error GoTo l_err
    
    Dim r As Range
    Dim i, n As Integer, beginRow As Integer
    Dim cols As String
    Dim arr() As String, colTgt As String
    
    If Selection.Columns.Count > 1 Then
        MsgBox "ѡ���������������У�"
        Exit Sub
    End If

    cols = Application.InputBox(prompt:="����Ҫ�ϲ�������(�ö��Ÿ������� E,F,G):", Type:=2,Default:="I,J,K,M,N,O")
    arr = Split(cols, ",")
    
    Application.DisplayAlerts = False
    
    For Each r In Selection
        If r.MergeCells Then
            If r.MergeArea.Columns.Count = 1 Then   '�ϲ����򣺵���
                If r.MergeArea.Cells.Offset.Address = r.Address Then
                    n = r.MergeArea.Count
                    beginRow = r.MergeArea.Cells.Offset.Row
                    
                    For i = 0 To UBound(arr)
                        colTgt = arr(i)
                        Range(colTgt & beginRow & ":" & colTgt & CStr(beginRow + n - 1)).Merge
                    Next i
                End If
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
    Exit Sub
l_err:
    Application.DisplayAlerts = True
    MsgBox "��������" & Err.Description
End Sub










