'���=�ظ�ֵ�����ֵ
'˵��=��˵��
Sub ͳ��һ����ÿ��ֵ���ֵĴ���()
        
    On Error Resume Next
    Dim rn As Range, ik, iv
    Dim count As Integer
    Dim d As Object, drow As Object
    
    Dim tar As Range
    Dim colName
    
    If Selection.Columns.count > 1 Then
        MsgBox "ѡ��ֻ�������һ���У�"
        Exit Sub
    End If
    '-------------------------------
    Set d = CreateObject("scripting.dictionary")
    Set drow = CreateObject("scripting.dictionary")
    
    For Each rn In Selection
        If rn <> "" Then
            If Not d.exists(rn.Value) Then
                d.Add rn.Value, 1
                drow.Add rn.Value, rn.Row
            Else
                count = d(rn.Value)
                d(rn.Value) = count + 1
            End If
        End If
    Next
    
    '-------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ��һ�հ������ڴ�Ž����", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    '---------------------
    colName = tar.Column
    
    ik = drow.keys
    iv = drow.items
    '---------------------
    For i = 0 To UBound(ik)
        Cells(iv(i), colName) = d(ik(i))
    Next
End Sub




