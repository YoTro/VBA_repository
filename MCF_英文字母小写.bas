'���=��ֵת��
'˵��=Ӣ����ĸСд

Sub Ӣ����ĸСд()
        
    On Error Resume Next
    Dim rn As Range
    Dim rlt

    For Each rn In Selection
    rlt = LCase(rn.Value)
    rn.Value = rlt
    Next
End Sub



