'���=���˳���
'˵��=��˵��
Sub ���һ���ַ���Ϊ�±�()
    On Error Resume Next
    Dim r As Range
    If Selection.Cells.Count >= 65536 Then
        MsgBox "ѡ�������̫���ˣ�����65536����Ԫ��"
        Exit sub
    End If
    
    Dim tar As Range
    Set tar = Selection
    Dim n As Integer
    For Each r In tar.Cells
        If r.Value <> "" Then
             n = r.Characters.Count
        
            With r.Characters(Start:=n, Length:=1).Font
                .Superscript = False
                .Subscript = True
            End With
        End If
    Next

End Sub
