'���=���˳���
'˵��=�����հ׵�Ԫ������ͣ������ŵ����Ϸ�
Public Sub ���հ����_��������Ϸ�()
        On Error Resume Next
        Dim all As Range
        Set all = Selection
        'Set all = Intersect(all, all.Worksheet.UsedRange)
        If all.Columns.Count > 65536 Or all.Rows.Count > 65536 Then
            MsgBox "ѡ��̫���ˣ����ж����ܳ���65536��", MsgBoxStyle.Information
            Return
        End If
        '------------------
        Dim r As Range
        Dim tmpr As Excel.Range
        Dim b As Integer
        Dim e As Integer
        Dim tmp As String
        For Each r In all.Columns
            b = 1
            e = r.Rows.Count
            For i = r.Rows.Count - 1 To 1 Step -1
                '------------------
                tmp = r.Cells(i, 1)
                If tmp = "" Then
                    b = i + 1
                    If b <= e Then
                        Set tmpr = r.Worksheet.Range(r.Cells(e, 1), r.Cells(b, 1))
                        r.Cells(i, 1).Formula = "=Sum( " & tmpr.Address & " )"
                    End If
                    e = i - 1
                End If
                '------------------
            Next
        Next

        MsgBox "���"
    End Sub
