'���=�ظ�ֵ�����ֵ
'˵��=��˵��
Sub ͳ��ѡ����ÿ��ֵ���ֵĴ���()
        
    On Error Resume Next
    Dim rn As Range
    Dim count As Integer
    Dim d As Object
    
    Dim tar As Range
    '-------------------------------
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
    
    '-------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ��һ��Ԫ�����ڴ�Ž����", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    '---------------------
    tar.Cells(1, 1).Offset(0, 0).Resize(d.count) = WorksheetFunction.Transpose(d.keys)
    tar.Cells(1, 1).Offset(0, 1).Resize(d.count) = WorksheetFunction.Transpose(d.items)
End Sub

