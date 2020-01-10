'类别=身份证工具
'说明=无说明
Sub 提取身份证的年龄()
    On Error Resume Next
    Dim ar, i, ii
    Dim tmp
    
    If Selection.Areas.Count > 1 Then Exit Sub
    If Selection.Cells.Count > Columns.Count Then
        MsgBox "您选择的区域过大！"
        Exit Sub
    End If

    ar = Selection
    Set rngs = Application.InputBox("请选择存放结果的区域", "提示", , , , , , 8)
    
    '一个单元格
    If Selection.Cells.Count = 1 Then
        tmp = IDAge(ar)
        ar = tmp
        
        rngs.Cells(1, 1) = ar
        Exit Sub
    End If
    
    '多个单元格
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tmp = IDAge(ar(i, ii))
            ar(i, ii) = tmp
        Next
    Next
    rngs.Resize(UBound(ar), UBound(ar, 2)) = ar
End Sub




Function IDAge(sid) As String
    Dim rlt As Date

    Select Case Len(sid)
        Case 15
            rlt = Format("19" & mid(sid, 7, 6), "0000-00-00")
        Case 18
            rlt = Format(mid(sid, 7, 8), "0000-00-00")
        Case 0
            IDAge = ""
            Exit Function
        Case Else
            IDAge = "无效"
            Exit Function
    End Select

    IDAge = Year(Date) - Year(rlt)
End Function






