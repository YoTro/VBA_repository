'���=
'˵��=��˵��

Sub ÿ����Ԫ���������()
    On Error Resume Next
    
    Dim rn As Range, x As Integer
    Dim max, min, unit As Double
    Dim arr() As Variant
    arr = Array("����", "����", "����")
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

Sub ÿ���ַ��������()
    On Error Resume Next
    
    Dim rn As Range, x As Integer
    Dim max, min, unit As Double
    Dim arr() As Variant
    arr = Array("����", "����", "����")
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


