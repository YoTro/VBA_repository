'���=��λ����
'˵��=ѡ����Χ��С��0�ĵ�Ԫ

Sub ѡ����Χ��С��0�ĵ�Ԫ()
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




