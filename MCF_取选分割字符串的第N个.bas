'类别=个人常用
'说明=
Option Explicit


Sub 取选分割字符串的第N个()
        
    On Error Resume Next
    Dim arr, tmp, rlt
    Dim x, count
    Dim pos As Integer
    Dim sepor, str
    Dim i
    Dim tar As Range
    
    If Selection.Columns.count > 1 Then Exit Sub
    '----------------------------------------------------
    sepor = Application.InputBox("请输入以什么作为分割字符串:", "输入", " ")
    If sepor = False Then Exit Sub
    
    str = Application.InputBox("请输入提取第几个:", "输入", "1")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    
    
    pos = CInt(str) - 1
    arr = Selection
    ReDim rlt(1 To UBound(arr))
    
    For i = 1 To UBound(arr)
        tmp = arr(i, 1)
        tmp = regReplace(tmp, "(" & sepor & ")+", sepor)
        
        x = Split(tmp, sepor)
        
        If UBound(x) >= pos Then
            rlt(i) = x(pos)
        Else
            rlt(i) = arr(i, 1)
        End If
        
    Next i

    '------------------------------------------------
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(按列)。", Title:="结果存放", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(UBound(rlt)) = WorksheetFunction.Transpose(rlt)
End Sub





Function regReplace(ByVal v As String, ByVal reg As String, ByVal rp As String) As String
    Dim rlt As String
    
    Dim re As Object
    Dim mhs
    Dim mh
    
    Set re = CreateObject("VBScript.regexp") ' regexp  '必须new
    re.Pattern = reg
    're.IgnoreCase = True
    re.Global = True  '多次匹配时需要
    
    rlt = re.Replace(v, rp)  '替换
    
    regReplace = rlt
End Function

