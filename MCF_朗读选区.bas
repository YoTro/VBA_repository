'类别=
'说明=朗读选区，请按ESC键终止
Sub 朗读选区()
    On Error Resume Next
    Dim r As Range
    Set r = Intersect(Selection, ActiveSheet.UsedRange) 'ActiveSheet 不能缺少
    
    Selection.Speak

End Sub


