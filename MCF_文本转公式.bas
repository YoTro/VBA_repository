'类别=数值转换
'说明=文本转公式
Sub 文本转公式()
        
    On Error Resume Next
    Dim rn As Range
    
    For Each rn In Selection
    rn.Formula = "=" & rn.Value
    Next
End Sub

