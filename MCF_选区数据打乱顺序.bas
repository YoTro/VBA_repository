'类别=重复值和随机值
'说明=选区数据打乱顺序


Sub 选区数据打乱顺序()
    Dim ar, i, ii
    Dim tmp, tr, tc
    
    If Selection.Areas.count > 1 Then Exit Sub
    If Selection.Cells.count > Columns.count Then
        MsgBox "您选择的区域过大！"
        Exit Sub
    End If
    
    ar = Selection
    
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tr = Int(Rnd * UBound(ar) + 1)
            tc = Int(Rnd * UBound(ar, 2) + 1)
            
            tmp = ar(tr, tc)
            ar(tr, tc) = ar(i, ii)
            ar(i, ii) = tmp
        Next
    Next
    

    Selection = ar
End Sub








