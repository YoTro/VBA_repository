'类别=数值转换
'说明=英文字母大写

Sub 英文字母大写()
        
    On Error Resume Next
    Dim rn As Range
    Dim rlt

    For Each rn In Selection
    rlt = UCase(rn.Value)
    rn.Value = rlt
    Next
End Sub





