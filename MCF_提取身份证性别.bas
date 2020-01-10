'类别=身份证工具
'说明=无说明
Sub 提取身份证性别()
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
        tmp = IDSex(ar)
        ar = tmp
        
        rngs.Cells(1, 1) = ar
        Exit Sub
    End If
    
    '多个单元格
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tmp = IDSex(ar(i, ii))
            ar(i, ii) = tmp
        Next
    Next
    rngs.Resize(UBound(ar), UBound(ar, 2)) = ar

End Sub





Function IDSex(sid)
    Dim s As String
    Select Case Len(sid)
        Case 15
            s = Right(sid, 1)
        Case 18
            s = mid(sid, 17, 1)
        Case 0
            IDSex = ""
            Exit Function
        Case Else
            IDSex = "无效身份证号"
            Exit Function
    End Select
    
    
    If Int(s / 2) = s / 2 Then              '是否为偶数
        IDSex = "女"                          '如果是，则性别=女
    Else                                    '否则
        IDSex = "男"                          '性别=女
    End If
End Function                                '结束循环








