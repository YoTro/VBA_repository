'类别=定位引用
'说明=选定范围内小于0的单元

Sub 选定范围内小于0的单元()
    Dim rng As Range
    Dim yvhf As String
    For Each rng In Selection
        If rng < 0 Then
            yvhf = yvhf & rng.Address & ","
        End If
    Next
    
    If yvhf <> "" Then
        Range(Left(yvhf, Len(yvhf) - 1)).Select
    End If
End Sub




