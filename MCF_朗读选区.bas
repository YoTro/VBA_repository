'���=
'˵��=�ʶ�ѡ�����밴ESC����ֹ
Sub �ʶ�ѡ��()
    On Error Resume Next
    Dim r As Range
    Set r = Intersect(Selection, ActiveSheet.UsedRange) 'ActiveSheet ����ȱ��
    
    Selection.Speak

End Sub


