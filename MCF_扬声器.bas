'类别=
'说明=无说明
Private Declare Function beep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
 
Sub 扬声器()
    '播放频率为 45000赫兹的扬声器堀音,持续 100微秒
    For I = 0 To 5
        beep 45000, 100
        DoEvents
    Next
End Sub


