'���=���֤����
'˵��=��˵��
Sub ��ȡ���֤�Ա�()
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
        tmp = IDSex(ar)
        ar = tmp
        
        rngs.Cells(1, 1) = ar
        Exit Sub
    End If
    
    '�����Ԫ��
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
            IDSex = "��Ч���֤��"
            Exit Function
    End Select
    
    
    If Int(s / 2) = s / 2 Then              '�Ƿ�Ϊż��
        IDSex = "Ů"                          '����ǣ����Ա�=Ů
    Else                                    '����
        IDSex = "��"                          '�Ա�=Ů
    End If
End Function                                '����ѭ��








