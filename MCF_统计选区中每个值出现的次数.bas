'类别=重复值和随机值
'说明=无说明
Sub 统计选区中每个值出现的次数()
        
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
    Set tar = Application.InputBox(prompt:="请选择一单元格用于存放结果。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    '---------------------
    tar.Cells(1, 1).Offset(0, 0).Resize(d.count) = WorksheetFunction.Transpose(d.keys)
    tar.Cells(1, 1).Offset(0, 1).Resize(d.count) = WorksheetFunction.Transpose(d.items)
End Sub

