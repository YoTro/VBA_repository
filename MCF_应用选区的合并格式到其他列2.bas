'���=�ϲ��Ͳ��
'˵��=��ѡ���еĺϲ���ʽ�ϲ������У����ݺϲ��������з�
Sub Ӧ��ѡ���ĺϲ���ʽ��������2()
    On Error GoTo l_err
    
    Dim r As Range, tmpr As Range
    Dim i, n As Integer, beginRow As Integer
    Dim cols As String
    Dim arr() As String, colTgt As String
    
    If Selection.Columns.Count > 1 Then
        MsgBox "ѡ���������������У�"
        Exit Sub
    End If

    cols = Application.InputBox(prompt:="����Ҫ�ϲ�������(�ö��Ÿ������� E,F,G):", Type:=2, Default:="I,J,K,M,N,O")
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
                        Set tmpr = Range(colTgt & beginRow & ":" & colTgt & CStr(beginRow + n - 1))
                        Dim x As String, x1 As String
						x=""
                        For Each t1 In tmpr.Cells
                            x1 = t1.Value
                            If x1 <> "" Then
                                x =IIf(x = "", x1, x & x1) ' IIf(x = "", x1, x & vbCrLf & x1)
                            End If
                        Next
                        tmpr.Merge
                        tmpr.Value = x
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

