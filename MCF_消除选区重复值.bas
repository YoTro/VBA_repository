'类别=重复值和随机值
'说明=在选区提取出不重复的序列
Sub 消除选区重复值()
        
    On Error Resume Next
    Dim rn As Range, res
    Dim tar
    
    Set d = CreateObject("scripting.dictionary")
    For Each rn In Selection
    If rn <> "" And Not d.exists(rn.Value) Then d.Add rn.Value, ""
    Next
    res = d.keys
    
    'For i = 0 To d.Count - 1
    	'Cells(i + 1, 5) = res(i)
    'Next
    
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(存放不重复序列,按列)。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(d.Count) = WorksheetFunction.Transpose(d.keys)
    'Cells(1, 11).Resize(d.Count) = WorksheetFunction.Transpose(d.keys)
End Sub

'[A:A].AdvancedFilter 2, , [e1], 1
'Selection.AdvancedFilter 2, , [s1], 1







