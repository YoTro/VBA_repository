'���=����ɾ��
'˵��=ɾ�����пհ׹�����
Sub ɾ���հ׹�����()
    Dim sht As Worksheet, n As Integer, iFlag As Boolean

    If MsgBox("Σ�ղ�����ȷ��ɾ����", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        'If sht.UsedRange.Cells.Count = 0 Then
        If Application.WorksheetFunction.CountA(sht.UsedRange.Cells) = 0 Then
            sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub



