'���=
'˵��=��˵��
Private Declare Function beep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
 
Sub ������()
    '����Ƶ��Ϊ 45000���ȵ�������ܥ��,���� 100΢��
    For I = 0 To 5
        beep 45000, 100
        DoEvents
    Next
End Sub


