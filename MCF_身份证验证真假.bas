'���=���֤����
'˵��=��˵��
Sub ���֤��֤���()
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
        tmp = CheckID(ar)
        ar = tmp
        
        rngs.Cells(1, 1) = ar
        Exit Sub
    End If
    
    '�����Ԫ��
    Randomize Timer
    For i = 1 To UBound(ar)
        For ii = 1 To UBound(ar, 2)
            tmp = CheckID(ar(i, ii))
            ar(i, ii) = tmp
        Next
    Next
    rngs.Resize(UBound(ar), UBound(ar, 2)) = ar

End Sub



Public Function CheckID(ByVal ID18 As String) As String
        Dim rlt As String
        Dim Ai(17) As Integer
        
        Select Case Len(ID18)
            Case 15
                CheckID = "�����֤��"
                Exit Function
            Case 18
            
            Case 0
                CheckID = ""
                Exit Function
            Case Else
                CheckID = "��Ч���֤��"
                Exit Function
        End Select
    

        
        CC = "10X98765432"
        Wi = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
        s = 0
        For i = 0 To 16
            Ai(i) = CInt(mid(ID18, i + 1, 1))
            s = s + Ai(i) * Wi(i)
        Next i
        rlt = mid(CC, s Mod 11 + 1, 1)
        
        If Right(ID18, 1) = rlt Then
            CheckID = "��"
        Else
            CheckID = "��"
        End If
End Function








