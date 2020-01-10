'类别=
'说明=无说明

Sub 每个单元格随机字体()
    On Error Resume Next
    
    Dim rn As Range, x As Integer
    Dim max, min, unit As Double
    Dim arr() As Variant
    arr = Array("宋体", "黑体", "隶书")
    unit = 1
    min = 0
    max = UBound(arr, 1)

    '--------------
    Randomize Timer
    Application.ScreenUpdating = False
    For Each rn In Selection
        x = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
        rn.Font.Name = arr(x)
    Next
    Application.ScreenUpdating = True
End Sub

Sub 每个字符随机字体()
    On Error Resume Next
    
    Dim rn As Range, x As Integer
    Dim max, min, unit As Double
    Dim arr() As Variant
    arr = Array("宋体", "黑体", "隶书")
    unit = 1
    min = 0
    max = UBound(arr, 1)

    '--------------
    Randomize Timer
    Application.ScreenUpdating = False
    For Each rn In Selection
        Dim maxLen
        maxLen = rn.Characters.count
        For j = 1 To maxLen - 1
            x = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
            rn.Characters(Start:=j, Length:=1).Font.Name = arr(x)
        Next
        
        'rn.Font.Name = arr(x)
    Next
    Application.ScreenUpdating = True
End Sub


