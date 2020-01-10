'类别=重复值和随机值
'说明=取选区出现一次的值

Sub 取选区出现一次的值()
        
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
    
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(存放不重复序列,按列)。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(d.count) = WorksheetFunction.Transpose(d.keys)
End Sub








