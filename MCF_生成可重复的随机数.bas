'类别=重复值和随机值
'说明=可重复的随机数生成
Sub 生成可重复的随机数()
    On Error Resume Next
    
    Dim count As Long
    Dim rn As Range
    Dim max, min, unit As Double
    Dim bRepeat As Boolean
    Dim d

    count = Selection.Cells.count   '16384列
    If count >= Columns.count Then
        MsgBox "请不要选择太大的区域！"
        Exit Sub
    End If

    min = Application.InputBox(prompt:="随机最小数字", Type:=1, Default:="0")
    If Not IsNumeric(min) Then Exit Sub    'false  不够准确，0也是false
    
    max = Application.InputBox(prompt:="随机最大数字", Type:=1, Default:="100")
    If Not IsNumeric(max) Then Exit Sub
    
    unit = Application.InputBox(prompt:="随机数的精确单位,如精确到1、精确到0.2 等等", Type:=1, Default:="1")
    If Not IsNumeric(unit) Then Exit Sub  'If unit = False Then Exit Sub
    
    Randomize Timer
    For Each rn In Selection
        rn = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
    Next

End Sub


