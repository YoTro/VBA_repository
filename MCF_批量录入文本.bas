'类别=个人常用
'说明=批量录入文本
Sub 批量录入文本()
    Dim r As Range
    Dim str
    str = Application.InputBox("请输入文本内容:", "输入文本内容")
    
    If str = False Then Exit Sub
    
    For Each r In Selection
        r =  str
    Next
End Sub





