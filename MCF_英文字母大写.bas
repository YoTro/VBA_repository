'���=��ֵת��
'˵��=Ӣ����ĸ��д

Sub Ӣ����ĸ��д()
        
    On Error Resume Next
    Dim rn As Range
    Dim rlt

    For Each rn In Selection
    rlt = UCase(rn.Value)
    rn.Value = rlt
    Next
End Sub





