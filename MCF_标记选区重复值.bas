'类别=重复值和随机值
'说明=将选区中的重复值染色

Sub 标记选区重复值()
    On Error Resume Next
    Dim rn As Range, first As Range
    Dim ColorIdx As Integer
    
    Set d = CreateObject("scripting.dictionary")
    Selection.Interior.ColorIndex = 2
    
    ColorIdx = 0
    For Each rn In Selection
        If rn <> "" Then
            If d.exists(rn.Value) Then
                Set first = Range(d(rn.Value))  '第一次出现的单元格
                If first.Interior.ColorIndex = 2 Then  '第一次出现时 未设置过颜色
                    '----------------------------------
                    ColorIdx = (ColorIdx + 1) Mod 56 + 1  '颜色可选范围：0~56
                    If ColorIdx = 2 Then ColorIdx = 3
                    '----------------------------------
                    first.Interior.ColorIndex = ColorIdx
                Else
                    ColorIdx = first.Interior.ColorIndex
                End If
                rn.Interior.ColorIndex = ColorIdx
            Else
                d.Add rn.Value, rn.Address
            End If
        End If
    Next

End Sub




