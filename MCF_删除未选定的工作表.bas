'���=����ɾ��
'˵��=ɾ��δѡ�������й�����

Sub ɾ��δѡ���Ĺ�����()
    Dim sht As Worksheet, n As Integer, iFlag As Boolean
    Dim ShtName() As String
    If MsgBox("Σ�ղ�����ȷ��ɾ����", vbOKCancel, "ע��!") = vbCancel Then
        Exit Sub
    End If

    n = ActiveWindow.SelectedSheets.Count
    ReDim ShtName(1 To n)
    n = 1
    For Each sht In ActiveWindow.SelectedSheets
        ShtName(n) = sht.Name
        n = n + 1
    Next
    Application.DisplayAlerts = False
    For Each sht In Sheets
        iFlag = False
        For i = 1 To n - 1
            If ShtName(i) = sht.Name Then
                iFlag = True
                Exit For
            End If
        Next
        If Not iFlag Then sht.Delete
    Next
    Application.DisplayAlerts = True
End Sub








