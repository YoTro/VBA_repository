'���=��ֵת��
'˵��=�ı�ת��ʽ
Sub �ı�ת��ʽ()
        
    On Error Resume Next
    Dim rn As Range
    
    For Each rn In Selection
    rn.Formula = "=" & rn.Value
    Next
End Sub

