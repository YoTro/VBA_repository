'���=���֤����
'˵��=��˵��
Sub ��ȡ���֤��������()
        On Error Resume Next
    Dim ar, i, ii
    Dim tmp
    
    If Selection.Areas.Count > 1 Then Exit Sub
    If Selection.Cells.Count > Columns.Count Then
        MsgBox "��ѡ����������"
        Exit Sub
    End If

    ar = Selection
    Set rngs = Application.InputBox("��ѡ���Ž��������", "��ʾ", , , , , , 8)
    
    'һ����Ԫ��
    If Selection.Cells.Count = 1 Then
        tmp = IDBirthday(ar)
        ar = tmp
        
        rngs.Cells(1, 1) = ar
        Exit Sub
    End If
    
    '�����Ԫ��
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tmp = IDBirthday(ar(i, ii))
            ar(i, ii) = tmp
        Next
    Next
    rngs.Resize(UBound(ar), UBound(ar, 2)) = ar


End Sub



Function IDBirthday(sid) As String
    Dim rlt

    Select Case Len(sid)
        Case 15
            rlt = Format("19" & mid(sid, 7, 6), "0000-00-00")
        Case 18
            rlt = Format(mid(sid, 7, 8), "0000-00-00")
        Case 0
            rlt = ""
        Case Else
            rlt = "��Ч"
    End Select

    IDBirthday = rlt
End Function






