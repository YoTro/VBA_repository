'���=�ظ�ֵ�����ֵ
'˵��=ȡѡ������һ�ε�ֵ

Sub ȡѡ������һ�ε�ֵ()
        
    On Error Resume Next
    Dim rn As Range, ik, iv
    Dim tar
    Dim count As Integer
    
    Set d = CreateObject("scripting.dictionary")
    For Each rn In Selection
        If rn <> "" Then
            If Not d.exists(rn.Value) Then
                d.Add rn.Value, 1
            Else
                count = d(rn.Value)
                d(rn.Value) = count + 1
            End If
        End If
    Next
    ik = d.keys
    iv = d.items
    
    For i = 0 To d.count - 1
        'MsgBox d(ik(i)) & "  " & ik(i) & ":" & iv(i)
        If d(ik(i)) > 1 Then
            d.Remove (ik(i))
        End If
    Next
    
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(��Ų��ظ�����,����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(d.count) = WorksheetFunction.Transpose(d.keys)
End Sub








