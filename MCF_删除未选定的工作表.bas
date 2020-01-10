'类别=批量删除
'说明=删除未选定的所有工作表

Sub 删除未选定的工作表()
    Dim sht As Worksheet, n As Integer, iFlag As Boolean
    Dim ShtName() As String
    If MsgBox("危险操作，确定删除？", vbOKCancel, "注意!") = vbCancel Then
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








