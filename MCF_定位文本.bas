'���=��λ����
'˵��=��ָ���ı���λ

Sub ��λ�ı�()
    Dim r As Range, a As Range
    Dim s
    s = Application.InputBox("������Ҫ��λ���ı�:", "����Ҫ��λ���ı�:", "�ϼ�")
    
    If s = False Then Exit Sub
    
    For Each a In ActiveSheet.UsedRange
        If a Like "*" & s & "*" Then
            If r Is Nothing Then
                Set r = a.Cells
            Else
                Set r = Union(r, a.Cells)
            End If
        End If
    Next
    
    If r Is Nothing Then
        MsgBox "δ�ҵ�ָ����Ԫ��!"
    Else
        r.Select
    End If
End Sub






