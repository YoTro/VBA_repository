'类别=重复值和随机值
'说明=生成不重复的随机数

Sub 生成不重复的随机数()
    On Error Resume Next
    
    Dim count As Long, needCount As Long
    Dim rn As Range
    Dim max, min, unit As Double
    Dim bRepeat As Boolean
    Dim d As Object
    Dim i, v

    count = Selection.Cells.count
    If count > 10000 Then
        MsgBox "请不要选择超过10000个单元格！"
        Exit Sub
    End If

    min = Application.InputBox(prompt:="随机最小数字", Type:=1, Default:="0")
    If Not IsNumeric(min) Then Exit Sub    'false  不够准确，0也是false
    
    max = Application.InputBox(prompt:="随机最大数字", Type:=1, Default:="100")
    If Not IsNumeric(max) Then Exit Sub
    
    unit = Application.InputBox(prompt:="随机数的精确单位,如精确到1、精确到0.2 等等", Type:=1, Default:="1")
    If Not IsNumeric(unit) Then Exit Sub  'If unit = False Then Exit Sub
    
    '---------------------------------------
    needCount = Int((max - min + unit) / unit)
    If count > needCount Then
        count = needCount  '合理个数
        'MsgBox "您选择的区域太大，无法生成不重复的随机数！ 至多只能选中" & needCount & "个单元格！"
        'Exit Sub
    End If
    
    Set d = CreateObject("scripting.dictionary")
    Randomize Timer
    For i = 1 To count
        Do
            v = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
        Loop While d.exists(v)

        Selection.Cells(i) = v
        d.Add v, ""
    Next i
End Sub





